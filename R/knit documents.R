#' A package to conveniently enable reproducible workflows for manuscripts
#'
#' @import officer
#' @import ggplot2
#' @import pandoc
#'
#' @rdname knit_docx
#' @title Knit placeholders to/from docx files
#' @description This function knits together an output .docx document from a template .docx document with placeholders and analysis code
#' @param template_docx_file The input .docx file containing the placeholders (in this case text in curly brackets{})
#' @param template_drive_ID The Google Drive ID of the template document containing the placeholders (in this case text in curly brackets{}). May be either a .docx or a Google Doc.
#' @param knitted_docx_file The output .docx file that knits together the figures/text/stats from the code
#' @param knitted_docx_google_ID The Google Drive ID of the output knit .docx file that will be overwritten on the knit. Note that this must point to an EXISTING .docx on google drive (can be a blank file uploaded for the purpose)
#' @param placeholder_generation_code_file (optional) If NA, the code will extract palceholder objects from the current global environment. Alternatively, this can be pointed to an existing .R script to run, runs the code, and will extract the placeholders from that run environment
#' @return Nothing
#' @export

knit_docx <- function(template_docx_file=NA,
                      template_drive_ID=NA,
                      knitted_docx_file = NA,
                      knitted_docx_google_ID=NA,
                      placeholder_generation_code_file = NA
                      ){
  # Libraries
  library(officer)
  library(ggplot2)
  library(pandoc)

  # Authorize google docs if needed
  if(!is.na(template_drive_ID) | !is.na(knitted_docx_google_ID)) {
    library(googledrive)
    drive_auth()
  }

  # Throw an error if file paths are not specified
  if (is.na(template_docx_file) & is.na(template_drive_ID)) {
    errorCondition("You must specify a path or google drive ID for the template doc.")
  }
  if (is.na(knitted_docx_file) & is.na(knitted_docx_google_ID)) {
    errorCondition("You must specify a path or google drive ID for the knitted doc.")
  }

  # Throw an error if BOTH file paths are not specified
  if (!is.na(template_docx_file) & !is.na(template_drive_ID)) {
    errorCondition("You must choose either a path or google drive ID for the template doc, not both.")
  }
  if (!is.na(knitted_docx_file) & !is.na(knitted_docx_google_ID)) {
    errorCondition("You must specify either a path or google drive ID for the knitted doc, not both.")
  }

  # Get all objects in the specified environment
  if (!is.na(placeholder_generation_code_file)){
    get_environment <- function(file){
      source(file,local=TRUE)
      return(rev(as.list(environment())))
    }
    generated_objects <- get_environment(placeholder_generation_code_file)
  } else {
    generated_objects <- rev(as.list(globalenv()))
  }

  # Get and prep the template doc
  docx_out <- tempfile(fileext = ".docx")
  if (!is.na(template_drive_ID)) {
    drive_file_name <- drive_get(as_id(template_drive_ID))$name
    if (substr(drive_file_name, nchar(drive_file_name)-4, nchar(drive_file_name))==".docx"){
      native.docx <- TRUE
    } else {native.docx <- FALSE}
    drive_download(file=as_id(template_drive_ID), path = docx_out)
  } else {
    native.docx <- TRUE
  }

  # Run docx-docx conversion hack to merge .docx chunks if native .docx.
  # This step is necessary because text blocks are stored in chunks, including
  # in some instances within the same word. To make things find/replaceable, the
  # full placeholder name must be all in one chunk. "Converting" the file reorganizes
  # the text chunks, allowing placeholders to be found/replaced consistently.
  if(native.docx == TRUE){
    pandoc::pandoc_convert(file = docx_out, from="docx", to = "docx",output=docx_out)
  }
  doc <- read_docx(path = docx_out)

  # Check classes of objects so they can be set in appropriate types
  object_names <- objects(generated_objects,sorted = FALSE)
  object_classes <- do.call(c,lapply(object_names,
                                     function(obj_name) { paste(class(generated_objects[[obj_name]]),collapse=",")
                                     }))

  text_object_class_blacklist <- c("function","gg","ggplot","ggplot,gg","gg,ggplot","bundled_ggplot")
  pontential_text_objects <- object_names[!object_classes %in% text_object_class_blacklist]

  pontential_figure_objects <- object_names[object_classes == "bundled_ggplot"]

  # Search the generated text placeholders and replace them if found in the template document
  for (obj_name in pontential_text_objects) {
    #print(obj_name)
    doc <- cursor_begin(doc)
    if (cursor_reach_test(doc, paste0("\\{",obj_name,"\\}"))){
      tryCatch(doc <- invisible(
        body_replace_all_text(
          doc,
          old_value = paste0("{",obj_name,"}"),
          new_value = as.character(generated_objects[[obj_name]]),
          #warn=FALSE,
          fixed=TRUE
        )
      ), error=function(e) {e}, warning=function(w) {w}
      )
    }
  }

  # Insert figures
  for (figure_name in pontential_figure_objects) {
    # Save ggplot object to a temporary file
    png_out <- tempfile(fileext = ".png")
    ggsave(filename = png_out,
           plot = generated_objects[[figure_name]]$plot,
           height = generated_objects[[figure_name]]$height,
           width = generated_objects[[figure_name]]$width,
           units = generated_objects[[figure_name]]$unit)

    # Find and replace figure placeholders
    if (cursor_reach_test(doc, paste0("\\{",figure_name,"\\}"))){
      doc <- cursor_begin(doc)
      doc <- cursor_reach(doc, paste0("\\{",figure_name,"\\}"))
      doc <- body_add_img(x = doc, src = png_out,pos="on",
                          height=generated_objects[[figure_name]]$height,
                          width=generated_objects[[figure_name]]$width,
                          unit=generated_objects[[figure_name]]$unit)
    }
  }

  # Export knitted doc
  if (!is.na(knitted_docx_file)) { print(doc, target=knitted_docx_file)}
  if (!is.na(knitted_docx_google_ID)) {
    temp_for_upload <- tempfile(fileext = ".docx")
    print(doc, target=temp_for_upload)
    drive_update(media=temp_for_upload,file=as_id(knitted_docx_google_ID))
    }
}

#' @examples
#' knit_docx(
#' template_docx_file = template_docx_file,
#' knitted_docx_file = knitted_docx_file)

#' @rdname bundle_ggplot
#' @title Bundles ggplot object with dimensions
#' @description A wrapper function for ggplot objects that specifies the dimensions they will be in. Needed because .docx and other document types require dimensions for images
#' @param plot A ggplot plot
#' @param width Numerical width of the ggplot # needed to insert into documents
#' @param height Numerical height of the ggplot # needed to insert into documents
#' @param unit The units the height and width are in (defaults to "in")
#' @export

# Bundle up ggplot objects so they can be inserted with dimensions into the document
bundle_ggplot <- function(plot,width=6.5,height=4,unit="in"){
  structure(list(plot=plot,height=height,width=width,unit=unit),class="bundled_ggplot")
}

#' @rdname preview_bundled_ggplot
#' @title Previews bundled ggplot object as it will appear in print
#' @description A convenience function so that you can preview what the bundle_ggplot object will look like in the specified dimensions
#' @param bundled_ggplot The bundled ggplot from the bundle_ggplot object
#' @export
# Quality of life function to show ggplot objects in their assigned scaling
# Code adapted from https://github.com/nflverse/nflplotR/blob/main/R/ggpreview.R

preview_bundled_ggplot <- function(bundled_ggplot){
  file <- tempfile()
  ggplot2::ggsave(
    file,
    plot = bundled_ggplot$plot,
    device = "png",
    scale = 1,
    width = bundled_ggplot$width,
    height = bundled_ggplot$height,
    units = bundled_ggplot$unit,
    limitsize = TRUE,
    bg = NULL
  )
  rstudioapi::viewer(file)
}

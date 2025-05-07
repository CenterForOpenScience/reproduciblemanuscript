# Libraries
library(Hmisc)
library(ggplot2)
library(reproduciblemanuscript) # devtools::install_github("noahhaber/reproduciblemanuscript")

# Placeholder statistics
start_date <- format(as.Date("05/02/2025","%m/%d/%Y"),format="%B %d, %Y")

n_journals <- 6
n_participants_returned_revisions <- 5
n_participants_editorial_decisions <- 6
n_authors_cons <- 525
n_participants_randomized <- 22
n_participants_RR <- 10
n_participants_SP <- 12
p_participants_editorial_decisions <- format.text.percent(n_participants_returned_revisions/n_participants_editorial_decisions)

# Example plot: Iris petals plot
{
  df.plot <- iris
  point.size <- 3
  p <- ggplot(data=df.plot,mapping=aes(x=Petal.Length,y=Petal.Width,color=Species))+
    geom_point(aes(color=Species),size=point.size)

  # Note: the plot must be bundled with width and height using the bundle_ggplot function as below
  figure_1 <- bundle_ggplot(p,height=10,width=10,unit="cm")

  preview_bundled_ggplot(figure_1)
}

# Run the knit
knit_docx(file_template_doc="template doc.docx",
                      file_knitted_doc="knitted doc.docx")

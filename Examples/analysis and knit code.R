# Libraries
  library(ggplot2)
  library(reproduciblemanuscript) # devtools::install_github("CenterForOpenScience/reproduciblemanuscript")

# Intro statistics
  start_date <- format(as.Date("05/02/2025","%m/%d/%Y"),format="%B %d, %Y")
  n_samples <- nrow(iris)
  n_species <- length(unique(iris$Species))

# Table 1
  table.1 <- do.call(rbind,lapply(levels(iris$Species),function(species) {
    data <- iris[iris$Species==species,]
    Sepal.Length <- paste0(round(mean(data$Sepal.Length),1), " (",
                           round(sd(data$Sepal.Length),1), ")")
    Sepal.Width <- paste0(round(mean(data$Sepal.Width),1), " (",
                           round(sd(data$Sepal.Width),1), ")")
    Petal.Length <- paste0(round(mean(data$Petal.Length),1), " (",
                           round(sd(data$Petal.Length),1), ")")
    Petal.Width <- paste0(round(mean(data$Petal.Width),1), " (",
                          round(sd(data$Petal.Width),1), ")")
    species <- paste0(toupper(substr(species, 1, 1)), substr(species, 2, nchar(species)))
    data.frame(species,Sepal.Length,Sepal.Width,Petal.Length,Petal.Width)
  }))

  for (i in seq_len(nrow(table.1))) {
    for (j in seq_len(ncol(table.1))) {
      var_name <- paste0("table_1_r", i, "_c", j)
      assign(var_name, table.1[i, j], envir = .GlobalEnv)
    }
  }

# Regression coefficients
  lm.fit <- lm(Petal.Length~Petal.Width,iris)
  lm.assoc.point.est <- round(summary(lm.fit)$coefficients[2,1],1)
  lm.assoc.CI.LB <- round(confint(lm.fit, level = 0.95)[2,1],1)
  lm.assoc.CI.UB <- round(confint(lm.fit, level = 0.95)[2,2],1)

# Figure 1
  df.plot <- iris
  point.size <- 3
  p <- ggplot(data=df.plot,mapping=aes(x=Petal.Length,y=Petal.Width,color=Species))+
    geom_point(aes(color=Species),size=point.size)
  # Note: the plot must be bundled with width and height using the bundle_ggplot function as below
  figure_1 <- bundle_ggplot(p,height=10,width=10,unit="cm")
  preview_bundled_ggplot(figure_1)

# Run the knit
  knit_docx(template_docx_file="template doc.docx",
            knitted_docx_file="knitted doc.docx")

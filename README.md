# Displaying and Visualising Data

Interactive Shiny app for the MPH *Fundamentals of Statistics* class demonstrations
(Community Hypertension Control Programme, n = 40).

- **Frequency distributions:** the data, frequency and grouped tables, categorical and
  quantitative displays, shape, misleading graphs.
- **Visualisation:** building a plot in layers, comparing groups (dynamite plot vs boxplot vs points),
  scatter plots and overplotting, when an axis may start above zero, colour and accessibility,
  small multiples, table or figure, spotting faults with the eight-point checklist, captions.
- **Build your slides:** load your own Excel or CSV file, confirm each variable's measurement scale,
  choose what to show and pick from the displays the app recommends (best, good, avoid, with reasons);
  edit the title, axis labels and caption, add each figure as a slide, and download the deck as
  PowerPoint (.pptx) or PDF.
- **Checkpoints:** ten scored questions with answers.

Every display has a PNG download and most show the ggplot2 code that reproduces them.

- `app.R`: the app (class dataset built in)
- `hypertension-dataset.xlsx`: the class dataset (sheets: Data, Dictionary); download/upload it from the app sidebar
- `manifest.json`: deployment manifest for Posit Connect Cloud

Run locally: `shiny::runApp()`

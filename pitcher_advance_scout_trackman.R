# an R script to help advance scout a pitcher using TrackMan data
# most everything is created in R but will need to make edits to the XLSX file before it is fully done
## these edits include: an overall summary of the pitcher, desciptions of his pitch shapes, and notes on his approach or any recent trends

library(readr)
library(openxlsx)
library(tidyverse)
library(ggpubr)
library(stringr)

# load data
trackman_data <- readr::read_csv('~/trackman_data.csv')

# filter data for specific pitcher
pitcher_data <- trackman_data %>% dplyr::filter(Pitcher == 'Austin, Kelly', TaggedPitchType != 'Other')

# get pitcher name in `First Last` string format
name_formatted <- stringr::str_glue(strsplit(pitcher_data$Pitcher[1], ', ')[[1]][2], strsplit(pitcher_data$Pitcher[1], ', ')[[1]][1], .sep = ' ')

# strike zone boundaries
strike_zone <- c(-.9, .9, 1.5, 3.5)

# is_in_zone: function used to determine number of pitches in strike zone
# param plate_z: vector of doubles representing a pitch's vertical location as it crosses home plate
# param plate_x: vector of doubles representing a pitch's horizontal location as it crosses home plate
# param zone: vector of strike zone boundaries (pitcher's right, pitcher's left, bottom, top)
is_in_zone <- function(plate_z, plate_x, zone = strike_zone) {
  zone <- 0
  for (i in seq_along(plate_z)) {
    if (is.na(plate_z[i]) | is.na(plate_x[i])) {
      zone <- zone
    } else if (between(plate_z[i], strike_zone[3], strike_zone[4]) &
               between(plate_x[i], strike_zone[1], strike_zone[2])) {
      zone <- zone + 1
    }
  }
  return(zone)
}

# coverting tagged pitch type to a factor so I can control the order
pitcher_data$TaggedPitchType <- factor(pitcher_data$TaggedPitchType, levels = c('Fastball', 'Sinker', 'Cutter', 'Curveball',
                                                                            'Slider', 'ChangeUp', 'Splitter', 'Other'))
# assigning colors to pitch types using Baseball Savant's coloring system
pitch_colors <- c('Fastball' = '#d22d49', 'Sinker' = '#fe9d00', 'Cutter' = '#933f2c', 'Curveball' = '#00d1ed',
                  'Slider' = '#c3bd0d', 'ChangeUp' = '#23be41', 'Splitter' = '#3bacac', 'Other' = '#Acafaf')

# creating `count` variable to be used for split count aggregation
pitcher_data$Count <- str_c(pitcher_data$Balls, '-', pitcher_data$Strikes)

# VS RHB
# creating tibble with pitcher's usage, zone rates, and velos (avg and high/low ranges) by pitch type vs RHB
pitcher_vr <- pitcher_data %>%
  dplyr::filter(BatterSide == 'Right') %>% 
  dplyr::group_by(TaggedPitchType) %>%
  dplyr::summarize(use = round(n()/sum(pitcher_data$BatterSide == 'Right'), 2) * 100,
                   zone = round(is_in_zone(PlateLocHeight, PlateLocSide)/n(), 2) * 100,
                   velo = paste0(format(round(mean(RelSpeed, na.rm = T), 1), nsmall = 1), ' (', round(quantile(RelSpeed, na.rm = T, probs = 0.1), 0), '-', round(quantile(RelSpeed, na.rm = T, probs = 0.9), 0), ')')) %>% 
  dplyr::arrange(desc(use)) %>% 
  tibble::add_column(description = NA) %>% 
  dplyr::rename(pitch = TaggedPitchType) %>% 
  dplyr::rename_with(str_to_upper)

# release plot vs RHB
vr_plot <- pitcher_data %>% 
  dplyr::filter(BatterSide == 'Right') %>% 
  dplyr::group_by(TaggedPitchType) %>% 
  dplyr::summarize(relz = mean(RelHeight, na.rm = T),
                   relx = mean(RelSide, na.rm = T)) %>% 
  dplyr::arrange(desc(TaggedPitchType)) %>% 
  ggplot2::ggplot(aes(x = -relx, y = relz, color = TaggedPitchType)) +
  geom_point(size = 3) +
  scale_color_manual(values = pitch_colors) +
  geom_point(x = -1.67, y = 5.83, color = 'black', size = 4.5) + # roughly repesents average RHP release point
  xlim(c(-3.5, 0.5)) +
  ylim(c(0, 7)) +
  theme_classic() +
  labs(x = '', y = '') +
  geom_segment(x = -.2, y = .3, xend = -.2, yend = .1, color = 'black') +
  geom_segment(x = .2, y = .3, xend = .2, yend = .1, color = 'black') +
  geom_segment(x = -.2, y = .1, xend = 0, yend = -.09, color = 'black') +
  geom_segment(x = .2, y = .1, xend = 0, yend = -.09, color = 'black') +
  geom_segment(x = -.2, y = .3, xend = .2, yend = .3, color = 'black') +
  theme(aspect.ratio = 1.5,
        legend.position = 'none',
        plot.margin = unit(c(0, -0.5, -0.5, -1), 'cm'))

# location plots
## generating tibble of pitcher's arsenal vs RHB (in order of usage)
vr_pitches <- pitcher_data %>%
  dplyr::filter(BatterSide == 'Right') %>% 
  dplyr::group_by(TaggedPitchType) %>% 
  dplyr::summarize(count = n()) %>% 
  dplyr::arrange(desc(count)) %>% 
  dplyr::select(TaggedPitchType)
# converting tibble to arsenal vector
vr_pitches_vec <- vr_pitches[[1]] %>% as.vector()

# usage/zone rate early count vs RHB
vr_early <- pitcher_data %>%
  dplyr::filter(Count %in% c('0-0', '1-0', '0-1', '1-1'), BatterSide == 'Right') %>% 
  dplyr::mutate(total = n()) %>% 
  dplyr::group_by(TaggedPitchType) %>%
  dplyr::summarize(use = round(n()/sum(mean(total)), 2) * 100,
                   zone = round(is_in_zone(PlateLocHeight, PlateLocSide)/n(), 2) * 100) %>% 
  dplyr::arrange(desc(use))

# initialize location plots list
early_pitch_plots_vr <- list()
for (pitch in vr_early$TaggedPitchType) {
  early_pitch_plots_vr[[which(vr_early$TaggedPitchType == pitch)]] <- pitcher_data %>% 
    dplyr::filter(TaggedPitchType == pitch, Count %in% c('0-0', '1-0', '0-1', '1-1'), BatterSide == 'Right') %>% 
    ggplot2::ggplot(aes(x = -PlateLocSide, y = PlateLocHeight, color = TaggedPitchType)) +
    geom_point(size = 2, alpha = 0.5) +
    # geom_mark_ellipse(aes(fill = TaggedPitchType)) +
    stat_ellipse(level = 0.5, lwd = 1.5) +            
    scale_color_manual(values = pitch_colors) +
    geom_segment(x = -.85, y = 1.5, xend = .85, yend = 1.5, color = 'black') +
    geom_segment(x = -.85, y = 3.3, xend = .85, yend = 3.3, color = 'black') +
    geom_segment(x = -.85, y = 1.5, xend = -.85, yend = 3.3, color = 'black') +
    geom_segment(x = .85, y = 1.5, xend = .85, yend = 3.3, color = 'black') +
    geom_segment(x = -.77, y = .5, xend = -.8, yend = .2, color = 'black') +
    geom_segment(x = .77, y = .5, xend = .8, yend = .2, color = 'black') +
    geom_segment(x = -.8, y = .2, xend = 0, yend = .01, color = 'black') +
    geom_segment(x = .8, y = .2, xend = 0, yend = .01, color = 'black') +
    geom_segment(x = -.77, y = .5, xend = .8, yend = .5, color = 'black') +
    labs(title = paste(pitch, vr_early[vr_early$TaggedPitchType == pitch, 2]), x = '', y = '', subtitle = vr_early[vr_early$TaggedPitchType == pitch, 3]) +
    theme_bw() +
    theme(plot.title = element_text(hjust = .95, vjust = -11, size = 16),
          plot.subtitle = element_text(hjust = 0.05, vjust = -53, size = 13),
          legend.position = 'none',
          panel.grid.major = element_blank(),
          panel.grid.minor = element_blank(),
          axis.ticks = element_blank(),
          axis.text.x = element_blank(),
          axis.text.y = element_blank(),
          plot.margin = unit(c(-2.5, -1, -1, -1), 'lines'),
          aspect.ratio = 1.5) +
    xlim(c(-1.5, 1.5)) +
    ylim(c(0, 4.25))
}
early_loc_vr <- ggpubr::ggarrange(plotlist = early_pitch_plots_vr, nrow = 1)  

# usage/zone rate hitters count vs RHB
vr_hitter <- pitcher_data %>%
  dplyr::filter(Count %in% c('2-0', '3-0', '2-1', '3-1'), BatterSide == 'Right') %>% 
  dplyr::mutate(total = n()) %>% 
  dplyr::group_by(TaggedPitchType) %>%
  dplyr::summarize(use = round(n()/sum(mean(total)), 2) * 100,
                   zone = round(is_in_zone(PlateLocHeight, PlateLocSide)/n(), 2) * 100) %>% 
  dplyr::arrange(desc(use))

hitter_pitch_plots_vr <- list()
for (pitch in vr_hitter$TaggedPitchType) {
  hitter_pitch_plots_vr[[which(vr_hitter$TaggedPitchType == pitch)]] <- pitcher_data %>% 
    dplyr::filter(TaggedPitchType == pitch, Count %in% c('2-0', '3-0', '2-1', '3-1'), BatterSide == 'Right') %>% 
    ggplot2::ggplot(aes(x = -PlateLocSide, y = PlateLocHeight, color = TaggedPitchType)) +
    geom_point(size = 2, alpha = 0.5) +
    # geom_mark_ellipse(aes(fill = TaggedPitchType)) +
    stat_ellipse(level = 0.5, lwd = 1.5) +            
    scale_color_manual(values = pitch_colors) +
    geom_segment(x = -.85, y = 1.5, xend = .85, yend = 1.5, color = 'black') +
    geom_segment(x = -.85, y = 3.3, xend = .85, yend = 3.3, color = 'black') +
    geom_segment(x = -.85, y = 1.5, xend = -.85, yend = 3.3, color = 'black') +
    geom_segment(x = .85, y = 1.5, xend = .85, yend = 3.3, color = 'black') +
    geom_segment(x = -.77, y = .5, xend = -.8, yend = .2, color = 'black') +
    geom_segment(x = .77, y = .5, xend = .8, yend = .2, color = 'black') +
    geom_segment(x = -.8, y = .2, xend = 0, yend = .01, color = 'black') +
    geom_segment(x = .8, y = .2, xend = 0, yend = .01, color = 'black') +
    geom_segment(x = -.77, y = .5, xend = .8, yend = .5, color = 'black') +
    labs(title = paste(pitch, vr_hitter[vr_hitter$TaggedPitchType == pitch, 2]), x = '', y = '', subtitle = vr_hitter[vr_hitter$TaggedPitchType == pitch, 3]) +
    theme_bw() +
    theme(plot.title = element_text(hjust = .95, vjust = -11, size = 16),
          plot.subtitle = element_text(hjust = 0.05, vjust = -53, size = 13),
          legend.position = 'none',
          panel.grid.major = element_blank(),
          panel.grid.minor = element_blank(),
          axis.ticks = element_blank(),
          axis.text.x = element_blank(),
          axis.text.y = element_blank(),
          plot.margin = unit(c(-2.5, -1, -1, -1), 'lines'),
          aspect.ratio = 1.5) +
    xlim(c(-1.5, 1.5)) +
    ylim(c(0, 4.25))
}
hitter_loc_vr <- ggpubr::ggarrange(plotlist = hitter_pitch_plots_vr, nrow = 1) 

# usage/zone rate finish count vs RHB
vr_finish <- pitcher_data %>%
  dplyr::filter(Count %in% c('0-2', '1-2', '2-2'), BatterSide == 'Right') %>% 
  dplyr::mutate(total = n()) %>% 
  dplyr::group_by(TaggedPitchType) %>%
  dplyr::summarize(use = round(n()/sum(mean(total)), 2) * 100,
                   zone = round(is_in_zone(PlateLocHeight, PlateLocSide)/n(), 2) * 100) %>% 
  dplyr::arrange(desc(use))

finish_pitch_plots_vr <- list()
for (pitch in vr_finish$TaggedPitchType) {
  finish_pitch_plots_vr[[which(vr_finish$TaggedPitchType == pitch)]] <- pitcher_data %>% 
    dplyr::filter(TaggedPitchType == pitch, Count %in% c('0-2', '1-2', '2-2'), BatterSide == 'Right') %>% 
    ggplot2::ggplot(aes(x = -PlateLocSide, y = PlateLocHeight, color = TaggedPitchType)) +
    geom_point(size = 2, alpha = 0.5) +
    # geom_mark_ellipse(aes(fill = TaggedPitchType)) +
    stat_ellipse(level = 0.5, lwd = 1.5) +            
    scale_color_manual(values = pitch_colors) +
    geom_segment(x = -.85, y = 1.5, xend = .85, yend = 1.5, color = 'black') +
    geom_segment(x = -.85, y = 3.3, xend = .85, yend = 3.3, color = 'black') +
    geom_segment(x = -.85, y = 1.5, xend = -.85, yend = 3.3, color = 'black') +
    geom_segment(x = .85, y = 1.5, xend = .85, yend = 3.3, color = 'black') +
    geom_segment(x = -.77, y = .5, xend = -.8, yend = .2, color = 'black') +
    geom_segment(x = .77, y = .5, xend = .8, yend = .2, color = 'black') +
    geom_segment(x = -.8, y = .2, xend = 0, yend = .01, color = 'black') +
    geom_segment(x = .8, y = .2, xend = 0, yend = .01, color = 'black') +
    geom_segment(x = -.77, y = .5, xend = .8, yend = .5, color = 'black') +
    labs(title = paste(pitch, vr_finish[vr_finish$TaggedPitchType == pitch, 2]), x = '', y = '', subtitle = vr_finish[vr_finish$TaggedPitchType == pitch, 3]) +
    theme_bw() +
    theme(plot.title = element_text(hjust = .95, vjust = -11, size = 16),
          plot.subtitle = element_text(hjust = 0.05, vjust = -53, size = 13),
          legend.position = 'none',
          panel.grid.major = element_blank(),
          panel.grid.minor = element_blank(),
          axis.ticks = element_blank(),
          axis.text.x = element_blank(),
          axis.text.y = element_blank(),
          plot.margin = unit(c(-2.5, -1, -1, -1), 'lines'),
          aspect.ratio = 1.5) +
    xlim(c(-1.5, 1.5)) +
    ylim(c(0, 4.25))
}
finish_loc_vr <- ggpubr::ggarrange(plotlist = finish_pitch_plots_vr, nrow = 1)

# usage/zone rate full count vs RHB
vr_full <- pitcher_data %>%
  dplyr::filter(Count == '3-2', BatterSide == 'Right') %>% 
  dplyr::mutate(total = n()) %>% 
  dplyr::group_by(TaggedPitchType) %>%
  dplyr::summarize(use = round(n()/sum(mean(total)), 2) * 100,
                   zone = round(is_in_zone(PlateLocHeight, PlateLocSide)/n(), 2) * 100) %>% 
  dplyr::arrange(desc(use))

full_pitch_plots_vr <- list()
for (pitch in vr_full$TaggedPitchType) {
  full_pitch_plots_vr[[which(vr_full$TaggedPitchType == pitch)]] <- pitcher_data %>% 
    dplyr::filter(TaggedPitchType == pitch, Count == '3-2', BatterSide == 'Right') %>% 
    ggplot2::ggplot(aes(x = -PlateLocSide, y = PlateLocHeight, color = TaggedPitchType)) +
    geom_point(size = 2, alpha = 0.5) +
    # geom_mark_ellipse(aes(fill = TaggedPitchType)) +
    stat_ellipse(level = 0.5, lwd = 1.5) +            
    scale_color_manual(values = pitch_colors) +
    geom_segment(x = -.85, y = 1.5, xend = .85, yend = 1.5, color = 'black') +
    geom_segment(x = -.85, y = 3.3, xend = .85, yend = 3.3, color = 'black') +
    geom_segment(x = -.85, y = 1.5, xend = -.85, yend = 3.3, color = 'black') +
    geom_segment(x = .85, y = 1.5, xend = .85, yend = 3.3, color = 'black') +
    geom_segment(x = -.77, y = .5, xend = -.8, yend = .2, color = 'black') +
    geom_segment(x = .77, y = .5, xend = .8, yend = .2, color = 'black') +
    geom_segment(x = -.8, y = .2, xend = 0, yend = .01, color = 'black') +
    geom_segment(x = .8, y = .2, xend = 0, yend = .01, color = 'black') +
    geom_segment(x = -.77, y = .5, xend = .8, yend = .5, color = 'black') +
    labs(title = paste(pitch, vr_full[vr_full$TaggedPitchType == pitch, 2]), x = '', y = '', subtitle = vr_full[vr_full$TaggedPitchType == pitch, 3]) +
    theme_bw() +
    theme(plot.title = element_text(hjust = .95, vjust = -11, size = 16),
          plot.subtitle = element_text(hjust = 0.05, vjust = -53, size = 13),
          legend.position = 'none',
          panel.grid.major = element_blank(),
          panel.grid.minor = element_blank(),
          axis.ticks = element_blank(),
          axis.text.x = element_blank(),
          axis.text.y = element_blank(),
          plot.margin = unit(c(-2.5, -1, -1, -1), 'lines'),
          aspect.ratio = 1.5) +
    xlim(c(-1.5, 1.5)) +
    ylim(c(0, 4.25))
}
full_loc_vr <- ggpubr::ggarrange(plotlist = full_pitch_plots_vr, nrow = 1)  


# title style
title_style <- openxlsx::createStyle(halign = 'center', valign = 'center', textDecoration = 'bold', fontSize = 36)
# subtitle style
subtitle_style <- openxlsx::createStyle(halign = 'center', valign = 'center', textDecoration = 'bold', fontSize = 18, fontColour = '#2774AE')
# section header style
subheader_style <- openxlsx::createStyle(halign = 'center', valign = 'center', textDecoration = c('bold', 'italic'), fontSize = 16,
                                         fgFill = '#2774AE', fontColour = 'white', border = c('top', 'bottom', 'left', 'right'), borderStyle = 'medium', borderColour = '#FFD100')
# data header style
header_style <- openxlsx::createStyle(halign = 'center', valign = 'center', textDecoration = 'bold',
                                      border = c('top', 'bottom', 'left', 'right'), borderStyle = 'medium', fontSize = 16)
# data style
data_style <- openxlsx::createStyle(halign = 'center', valign = 'center', textDecoration = 'bold', fontSize = 14,
                                    border = c('top', 'bottom', 'left', 'right'), borderStyle = 'medium')
# notes boxes style
notes_style <- openxlsx::createStyle(halign = 'center', valign = 'center', textDecoration = 'bold', fontSize = 16, fontColour = '#2774AE', wrapText = TRUE,
                                     border = c('top', 'bottom', 'left', 'right'), borderStyle = 'medium')
# loc plot title style
locheader_style <- openxlsx::createStyle(halign = 'center', valign = 'center', textRotation = 90, textDecoration = 'bold', fontSize = 16,
                                         fgFill = '#2774AE', fontColour = 'white', border = c('top', 'bottom', 'left', 'right'), borderStyle = 'medium', borderColour = '#FFD100')

# creating workbook
wb <- openxlsx::createWorkbook()
# create vs RHB sheet
addWorksheet(wb, paste(name_formatted, 'vs RHB'))

## some sheet formatting
pageSetup(wb, 1, orientation = 'portrait', left = 0.5, right = 0.5, top = 0.5, bottom = 0.5,
          fitToWidth = TRUE, fitToHeight = TRUE, paperSize = 1)

# merging cells
mergeCells(wb, 1, rows = 1, cols = 2:8)
mergeCells(wb, 1, rows = 2, cols = 1:9)
mergeCells(wb, 1, rows = 3, cols = 4:9)
mergeCells(wb, 1, rows = 4, cols = 8:9)
mergeCells(wb, 1, rows = 5, cols = 8:9)
mergeCells(wb, 1, rows = 6, cols = 8:9)
mergeCells(wb, 1, rows = 7, cols = 8:9)
mergeCells(wb, 1, rows = 8, cols = 8:9)
mergeCells(wb, 1, rows = 9, cols = 8:9)
mergeCells(wb, 1, rows = 10, cols = 1:9)
mergeCells(wb, 1, rows = 11, cols = 4:9)
mergeCells(wb, 1, rows = 12:20, cols = 1)
mergeCells(wb, 1, rows = 21:29, cols = 1)
mergeCells(wb, 1, rows = 30:38, cols = 1)
mergeCells(wb, 1, rows = 39:47, cols = 1)
mergeCells(wb, 1, rows = 12:22, cols = 3)
mergeCells(wb, 1, rows = 23:33, cols = 3)
mergeCells(wb, 1, rows = 34:44, cols = 3)
mergeCells(wb, 1, rows = 45:55, cols = 3)

# removing gridlines
# showGridLines(wb, 1, showGridLines = FALSE)

# setting row heights
setRowHeights(wb, 1, rows = 1, heights = 36)
setRowHeights(wb, 1, rows = 2, heights = 30)
setRowHeights(wb, 1, rows = 3, heights = 18)
setRowHeights(wb, 1, rows = 4, heights = 18)
setRowHeights(wb, 1, rows = 5:9, heights = 36)
setRowHeights(wb, 1, rows = 10, heights = 14)
setRowHeights(wb, 1, rows = 11, heights = 18)

# setting column widths
setColWidths(wb, 1, cols = 1, widths = 25)
setColWidths(wb, 1, cols = 2, widths = 2)
setColWidths(wb, 1, cols = 3, widths = 3)
setColWidths(wb, 1, cols = 4, widths = 15)
setColWidths(wb, 1, cols = 5:6, widths = 8)
setColWidths(wb, 1, cols = 7:9, widths = 18)

# inserting TM logo
insertImage(wb, 1, '~/Downloads/TM_logo.png', height = 0.42, width = 2.05, units = 'in', startRow = 1, startCol = 1)

# inserting opp logo
insertImage(wb, 1, '~/Downloads/ucla_opp_logo.png', height = 0.75, width = 1.5, units = 'in', startRow = 1, startCol = 9)

# inserting UCLA (team) logo
insertImage(wb, 1, '~/Downloads/ucla_logo.png', height = 2, width = 2.1, units = 'in', startRow = 49, startCol = 1)

# adding title to sheet
writeData(wb, 1, x = paste(name_formatted, 'vs RHB'), startRow = 1, startCol = 2)
addStyle(wb, 1, rows = 1, cols = 2, style = title_style)

# adding subtitle to sheet
addStyle(wb, 1, rows = 2, cols = 1, style = subtitle_style)

# inserting release plot
writeData(wb, 1, x = 'RELEASE', startRow = 3, startCol = 1)
addStyle(wb, 1, rows = 3, cols = 1, style = subheader_style)
print(vr_plot)
insertPlot(wb, 1, startRow = 4, height = 2.65, width = 2.1, units = 'in')

# inserting pitch table
writeData(wb, 1, x = 'ARSENAL', startRow = 3, startCol = 4)
addStyle(wb, 1, rows = 3, cols = 4:9, style = subheader_style, gridExpand = TRUE)
writeData(wb, 1, pitcher_vr, startRow = 4, startCol = 4)
addStyle(wb, 1, rows = 4, cols = 4:9, style = header_style)
addStyle(wb, 1, rows = which(grepl(pitcher_vr$PITCH, pattern = 'Fastball')) + 4, cols = 4:9,
         style = createStyle(fontColour = '#d22d49', fontSize = 16, textDecoration = 'bold', valign = 'center', halign = 'center',
                             border = c('top', 'bottom', 'left', 'right'), borderStyle = 'thin'))
addStyle(wb, 1, rows = which(grepl(pitcher_vr$PITCH, pattern = 'Sinker')) + 4, cols = 4:9,
         style = createStyle(fontColour = '#fe9d00', fontSize = 16, textDecoration = 'bold', valign = 'center', halign = 'center',
                             border = c('top', 'bottom', 'left', 'right'), borderStyle = 'thin'))
addStyle(wb, 1, rows = which(grepl(pitcher_vr$PITCH, pattern = 'Cutter')) + 4, cols = 4:9,
         style = createStyle(fontColour = '#933f2c', fontSize = 16, textDecoration = 'bold', valign = 'center', halign = 'center',
                             border = c('top', 'bottom', 'left', 'right'), borderStyle = 'thin'))
addStyle(wb, 1, rows = which(grepl(pitcher_vr$PITCH, pattern = 'Curveball')) + 4, cols = 4:9,
         style = createStyle(fontColour = '#00d1ed', fontSize = 16, textDecoration = 'bold', valign = 'center', halign = 'center',
                             border = c('top', 'bottom', 'left', 'right'), borderStyle = 'thin'))
addStyle(wb, 1, rows = which(grepl(pitcher_vr$PITCH, pattern = 'Slider')) + 4, cols = 4:9,
         style = createStyle(fontColour = '#c3bd0d', fontSize = 16, textDecoration = 'bold', valign = 'center', halign = 'center',
                             border = c('top', 'bottom', 'left', 'right'), borderStyle = 'thin'))
addStyle(wb, 1, rows = which(grepl(pitcher_vr$PITCH, pattern = 'ChangeUp')) + 4, cols = 4:9,
         style = createStyle(fontColour = '#23be41', fontSize = 16, textDecoration = 'bold', valign = 'center', halign = 'center',
                             border = c('top', 'bottom', 'left', 'right'), borderStyle = 'thin'))
addStyle(wb, 1, rows = which(grepl(pitcher_vr$PITCH, pattern = 'Splitter')) + 4, cols = 4:9,
         style = createStyle(fontColour = '#3bacac', fontSize = 16, textDecoration = 'bold', valign = 'center', halign = 'center',
                             border = c('top', 'bottom', 'left', 'right'), borderStyle = 'thin'))
addStyle(wb, 1, rows = which(grepl(pitcher_vr$PITCH, pattern = 'Other')) + 4, cols = 4:9,
         style = createStyle(fontColour = '#Acafaf', fontSize = 16, textDecoration = 'bold', valign = 'center', halign = 'center',
                             border = c('top', 'bottom', 'left', 'right'), borderStyle = 'thin'))

# notes section vs RHB
writeData(wb, 1, x = 'NOTES', startRow = 11, startCol = 1)
addStyle(wb, 1, rows = 11, cols = 1, style = subheader_style)
# writeData(wb, 1, 'test 1', startRow = 11, startCol = 1)
addStyle(wb, 1, rows = 12:20, cols = 1, style = notes_style)
# writeData(wb, 1, 'test 2', startRow = 20, startCol = 1)
addStyle(wb, 1, rows = 21:29, cols = 1, style = notes_style)
# writeData(wb, 1, 'test 3', startRow = 29, startCol = 1)
addStyle(wb, 1, rows = 30:38, cols = 1, style = notes_style)
# writeData(wb, 1, 'test 4', startRow = 38, startCol = 1)
addStyle(wb, 1, rows = 39:47, cols = 1, style = notes_style)

# location plots vs RHB
writeData(wb, 1, 'LOCATION', startRow = 11, startCol = 4)
addStyle(wb, 1, rows = 11, cols = 4:9, style = subheader_style)
writeData(wb, 1, 'EARLY', startRow = 12, startCol = 3)
addStyle(wb, 1, rows = 12:22, cols = 3, style = locheader_style)
print(early_loc_vr)
insertPlot(wb, 1, startRow = 12, startCol = 4, height = 2.15, width = 7.35, units = 'in')
writeData(wb, 1, 'HITTERS', startRow = 23, startCol = 3)
addStyle(wb, 1, rows = 23:33, cols = 3, style = locheader_style)
print(hitter_loc_vr)
insertPlot(wb, 1, startRow = 23, startCol = 4, height = 2.15, width = 7.35, units = 'in')
writeData(wb, 1, '2-STRIKE', startRow = 34, startCol = 3)
addStyle(wb, 1, rows = 34:44, cols = 3, style = locheader_style)
print(finish_loc_vr)
insertPlot(wb, 1, startRow = 34, startCol = 4, height = 2.15, width = 7.35, units = 'in')
writeData(wb, 1, 'FULL', startRow = 45, startCol = 3)
addStyle(wb, 1, rows = 45:55, cols = 3, style = locheader_style)
print(full_loc_vr)
insertPlot(wb, 1, startRow = 45, startCol = 4, height = 2.15, width = 7.35, units = 'in')

# VS LHB
# creating tibble with pitcher's usage, zone rates, and velos (avg and high/low ranges) by pitch type vs LHB
pitcher_vl <- pitcher_data %>%
  dplyr::filter(BatterSide == 'Left') %>% 
  dplyr::group_by(TaggedPitchType) %>%
  dplyr::summarize(use = round(n()/sum(pitcher_data$BatterSide == 'Left'), 2) * 100,
                   zone = round(is_in_zone(PlateLocHeight, PlateLocSide)/n(), 2) * 100,
                   velo = paste0(format(round(mean(RelSpeed, na.rm = T), 1), nsmall = 1), ' (', round(quantile(RelSpeed, na.rm = T, probs = 0.1), 0), '-', round(quantile(RelSpeed, na.rm = T, probs = 0.9), 0), ')')) %>% 
  dplyr::arrange(desc(use)) %>% 
  tibble::add_column(description = NA) %>% 
  dplyr::rename(pitch = TaggedPitchType) %>% 
  dplyr::rename_with(str_to_upper)

# release plot vs LHB
vr_plot <- pitcher_data %>% 
  dplyr::filter(BatterSide == 'Left') %>% 
  dplyr::group_by(TaggedPitchType) %>% 
  dplyr::summarize(relz = mean(RelHeight, na.rm = T),
                   relx = mean(RelSide, na.rm = T)) %>% 
  dplyr::arrange(desc(TaggedPitchType)) %>% 
  ggplot2::ggplot(aes(x = -relx, y = relz, color = TaggedPitchType)) +
  geom_point(size = 3) +
  scale_color_manual(values = pitch_colors) +
  geom_point(x = -1.67, y = 5.83, color = 'black', size = 4.5) + # roughly repesents average RHP release point
  xlim(c(-3.5, 0.5)) +
  ylim(c(0, 7)) +
  theme_classic() +
  labs(x = '', y = '') +
  geom_segment(x = -.2, y = .3, xend = -.2, yend = .1, color = 'black') +
  geom_segment(x = .2, y = .3, xend = .2, yend = .1, color = 'black') +
  geom_segment(x = -.2, y = .1, xend = 0, yend = -.09, color = 'black') +
  geom_segment(x = .2, y = .1, xend = 0, yend = -.09, color = 'black') +
  geom_segment(x = -.2, y = .3, xend = .2, yend = .3, color = 'black') +
  theme(aspect.ratio = 1.5,
        legend.position = 'none',
        plot.margin = unit(c(0, -0.5, -0.5, -1), 'cm'))

# location plots
## generating tibble of pitcher's arsenal vs LHB (in order of usage)
vr_pitches <- pitcher_data %>%
  dplyr::filter(BatterSide == 'Left') %>% 
  dplyr::group_by(TaggedPitchType) %>% 
  dplyr::summarize(count = n()) %>% 
  dplyr::arrange(desc(count)) %>% 
  dplyr::select(TaggedPitchType)
# converting tibble to arsenal vector
vr_pitches_vec <- vr_pitches[[1]] %>% as.vector()

# usage/zone rate early count vs LHB
vr_early <- pitcher_data %>%
  dplyr::filter(Count %in% c('0-0', '1-0', '0-1', '1-1'), BatterSide == 'Left') %>% 
  dplyr::mutate(total = n()) %>% 
  dplyr::group_by(TaggedPitchType) %>%
  dplyr::summarize(use = round(n()/sum(mean(total)), 2) * 100,
                   zone = round(is_in_zone(PlateLocHeight, PlateLocSide)/n(), 2) * 100) %>% 
  dplyr::arrange(desc(use))

# initialize location plots list
early_pitch_plots_vl <- list()
for (pitch in vr_early$TaggedPitchType) {
  early_pitch_plots_vl[[which(vr_early$TaggedPitchType == pitch)]] <- pitcher_data %>% 
    dplyr::filter(TaggedPitchType == pitch, Count %in% c('0-0', '1-0', '0-1', '1-1'), BatterSide == 'Left') %>% 
    ggplot2::ggplot(aes(x = -PlateLocSide, y = PlateLocHeight, color = TaggedPitchType)) +
    geom_point(size = 2, alpha = 0.5) +
    # geom_mark_ellipse(aes(fill = TaggedPitchType)) +
    stat_ellipse(level = 0.5, lwd = 1.5) +            
    scale_color_manual(values = pitch_colors) +
    geom_segment(x = -.85, y = 1.5, xend = .85, yend = 1.5, color = 'black') +
    geom_segment(x = -.85, y = 3.3, xend = .85, yend = 3.3, color = 'black') +
    geom_segment(x = -.85, y = 1.5, xend = -.85, yend = 3.3, color = 'black') +
    geom_segment(x = .85, y = 1.5, xend = .85, yend = 3.3, color = 'black') +
    geom_segment(x = -.77, y = .5, xend = -.8, yend = .2, color = 'black') +
    geom_segment(x = .77, y = .5, xend = .8, yend = .2, color = 'black') +
    geom_segment(x = -.8, y = .2, xend = 0, yend = .01, color = 'black') +
    geom_segment(x = .8, y = .2, xend = 0, yend = .01, color = 'black') +
    geom_segment(x = -.77, y = .5, xend = .8, yend = .5, color = 'black') +
    labs(title = paste(pitch, vr_early[vr_early$TaggedPitchType == pitch, 2]), x = '', y = '', subtitle = vr_early[vr_early$TaggedPitchType == pitch, 3]) +
    theme_bw() +
    theme(plot.title = element_text(hjust = .95, vjust = -11, size = 16),
          plot.subtitle = element_text(hjust = 0.05, vjust = -53, size = 13),
          legend.position = 'none',
          panel.grid.major = element_blank(),
          panel.grid.minor = element_blank(),
          axis.ticks = element_blank(),
          axis.text.x = element_blank(),
          axis.text.y = element_blank(),
          plot.margin = unit(c(-2.5, -1, -1, -1), 'lines'),
          aspect.ratio = 1.5) +
    xlim(c(-1.5, 1.5)) +
    ylim(c(0, 4.25))
}
early_loc_vl <- ggpubr::ggarrange(plotlist = early_pitch_plots_vl, nrow = 1)  

# usage/zone rate hitters count vs LHB
vr_hitter <- pitcher_data %>%
  dplyr::filter(Count %in% c('2-0', '3-0', '2-1', '3-1'), BatterSide == 'Left') %>% 
  dplyr::mutate(total = n()) %>% 
  dplyr::group_by(TaggedPitchType) %>%
  dplyr::summarize(use = round(n()/sum(mean(total)), 2) * 100,
                   zone = round(is_in_zone(PlateLocHeight, PlateLocSide)/n(), 2) * 100) %>% 
  dplyr::arrange(desc(use))

hitter_pitch_plots_vl <- list()
for (pitch in vr_hitter$TaggedPitchType) {
  hitter_pitch_plots_vl[[which(vr_hitter$TaggedPitchType == pitch)]] <- pitcher_data %>% 
    dplyr::filter(TaggedPitchType == pitch, Count %in% c('2-0', '3-0', '2-1', '3-1'), BatterSide == 'Left') %>% 
    ggplot2::ggplot(aes(x = -PlateLocSide, y = PlateLocHeight, color = TaggedPitchType)) +
    geom_point(size = 2, alpha = 0.5) +
    # geom_mark_ellipse(aes(fill = TaggedPitchType)) +
    stat_ellipse(level = 0.5, lwd = 1.5) +            
    scale_color_manual(values = pitch_colors) +
    geom_segment(x = -.85, y = 1.5, xend = .85, yend = 1.5, color = 'black') +
    geom_segment(x = -.85, y = 3.3, xend = .85, yend = 3.3, color = 'black') +
    geom_segment(x = -.85, y = 1.5, xend = -.85, yend = 3.3, color = 'black') +
    geom_segment(x = .85, y = 1.5, xend = .85, yend = 3.3, color = 'black') +
    geom_segment(x = -.77, y = .5, xend = -.8, yend = .2, color = 'black') +
    geom_segment(x = .77, y = .5, xend = .8, yend = .2, color = 'black') +
    geom_segment(x = -.8, y = .2, xend = 0, yend = .01, color = 'black') +
    geom_segment(x = .8, y = .2, xend = 0, yend = .01, color = 'black') +
    geom_segment(x = -.77, y = .5, xend = .8, yend = .5, color = 'black') +
    labs(title = paste(pitch, vr_hitter[vr_hitter$TaggedPitchType == pitch, 2]), x = '', y = '', subtitle = vr_hitter[vr_hitter$TaggedPitchType == pitch, 3]) +
    theme_bw() +
    theme(plot.title = element_text(hjust = .95, vjust = -11, size = 16),
          plot.subtitle = element_text(hjust = 0.05, vjust = -53, size = 13),
          legend.position = 'none',
          panel.grid.major = element_blank(),
          panel.grid.minor = element_blank(),
          axis.ticks = element_blank(),
          axis.text.x = element_blank(),
          axis.text.y = element_blank(),
          plot.margin = unit(c(-2.5, -1, -1, -1), 'lines'),
          aspect.ratio = 1.5) +
    xlim(c(-1.5, 1.5)) +
    ylim(c(0, 4.25))
}
hitter_loc_vl <- ggpubr::ggarrange(plotlist = hitter_pitch_plots_vl, nrow = 1) 

# usage/zone rate finish count vs LHB
vr_finish <- pitcher_data %>%
  dplyr::filter(Count %in% c('0-2', '1-2', '2-2'), BatterSide == 'Left') %>% 
  dplyr::mutate(total = n()) %>% 
  dplyr::group_by(TaggedPitchType) %>%
  dplyr::summarize(use = round(n()/sum(mean(total)), 2) * 100,
                   zone = round(is_in_zone(PlateLocHeight, PlateLocSide)/n(), 2) * 100) %>% 
  dplyr::arrange(desc(use))

finish_pitch_plots_vl <- list()
for (pitch in vr_finish$TaggedPitchType) {
  finish_pitch_plots_vl[[which(vr_finish$TaggedPitchType == pitch)]] <- pitcher_data %>% 
    dplyr::filter(TaggedPitchType == pitch, Count %in% c('0-2', '1-2', '2-2'), BatterSide == 'Left') %>% 
    ggplot2::ggplot(aes(x = -PlateLocSide, y = PlateLocHeight, color = TaggedPitchType)) +
    geom_point(size = 2, alpha = 0.5) +
    # geom_mark_ellipse(aes(fill = TaggedPitchType)) +
    stat_ellipse(level = 0.5, lwd = 1.5) +            
    scale_color_manual(values = pitch_colors) +
    geom_segment(x = -.85, y = 1.5, xend = .85, yend = 1.5, color = 'black') +
    geom_segment(x = -.85, y = 3.3, xend = .85, yend = 3.3, color = 'black') +
    geom_segment(x = -.85, y = 1.5, xend = -.85, yend = 3.3, color = 'black') +
    geom_segment(x = .85, y = 1.5, xend = .85, yend = 3.3, color = 'black') +
    geom_segment(x = -.77, y = .5, xend = -.8, yend = .2, color = 'black') +
    geom_segment(x = .77, y = .5, xend = .8, yend = .2, color = 'black') +
    geom_segment(x = -.8, y = .2, xend = 0, yend = .01, color = 'black') +
    geom_segment(x = .8, y = .2, xend = 0, yend = .01, color = 'black') +
    geom_segment(x = -.77, y = .5, xend = .8, yend = .5, color = 'black') +
    labs(title = paste(pitch, vr_finish[vr_finish$TaggedPitchType == pitch, 2]), x = '', y = '', subtitle = vr_finish[vr_finish$TaggedPitchType == pitch, 3]) +
    theme_bw() +
    theme(plot.title = element_text(hjust = .95, vjust = -11, size = 16),
          plot.subtitle = element_text(hjust = 0.05, vjust = -53, size = 13),
          legend.position = 'none',
          panel.grid.major = element_blank(),
          panel.grid.minor = element_blank(),
          axis.ticks = element_blank(),
          axis.text.x = element_blank(),
          axis.text.y = element_blank(),
          plot.margin = unit(c(-2.5, -1, -1, -1), 'lines'),
          aspect.ratio = 1.5) +
    xlim(c(-1.5, 1.5)) +
    ylim(c(0, 4.25))
}
finish_loc_vl <- ggpubr::ggarrange(plotlist = finish_pitch_plots_vl, nrow = 1)

# usage/zone rate full count vs LHB
vr_full <- pitcher_data %>%
  dplyr::filter(Count == '3-2', BatterSide == 'Left') %>% 
  dplyr::mutate(total = n()) %>% 
  dplyr::group_by(TaggedPitchType) %>%
  dplyr::summarize(use = round(n()/sum(mean(total)), 2) * 100,
                   zone = round(is_in_zone(PlateLocHeight, PlateLocSide)/n(), 2) * 100) %>% 
  dplyr::arrange(desc(use))

full_pitch_plots_vl <- list()
for (pitch in vr_full$TaggedPitchType) {
  full_pitch_plots_vl[[which(vr_full$TaggedPitchType == pitch)]] <- pitcher_data %>% 
    dplyr::filter(TaggedPitchType == pitch, Count == '3-2', BatterSide == 'Left') %>% 
    ggplot2::ggplot(aes(x = -PlateLocSide, y = PlateLocHeight, color = TaggedPitchType)) +
    geom_point(size = 2, alpha = 0.5) +
    # geom_mark_ellipse(aes(fill = TaggedPitchType)) +
    stat_ellipse(level = 0.5, lwd = 1.5) +            
    scale_color_manual(values = pitch_colors) +
    geom_segment(x = -.85, y = 1.5, xend = .85, yend = 1.5, color = 'black') +
    geom_segment(x = -.85, y = 3.3, xend = .85, yend = 3.3, color = 'black') +
    geom_segment(x = -.85, y = 1.5, xend = -.85, yend = 3.3, color = 'black') +
    geom_segment(x = .85, y = 1.5, xend = .85, yend = 3.3, color = 'black') +
    geom_segment(x = -.77, y = .5, xend = -.8, yend = .2, color = 'black') +
    geom_segment(x = .77, y = .5, xend = .8, yend = .2, color = 'black') +
    geom_segment(x = -.8, y = .2, xend = 0, yend = .01, color = 'black') +
    geom_segment(x = .8, y = .2, xend = 0, yend = .01, color = 'black') +
    geom_segment(x = -.77, y = .5, xend = .8, yend = .5, color = 'black') +
    labs(title = paste(pitch, vr_full[vr_full$TaggedPitchType == pitch, 2]), x = '', y = '', subtitle = vr_full[vr_full$TaggedPitchType == pitch, 3]) +
    theme_bw() +
    theme(plot.title = element_text(hjust = .95, vjust = -11, size = 16),
          plot.subtitle = element_text(hjust = 0.05, vjust = -53, size = 13),
          legend.position = 'none',
          panel.grid.major = element_blank(),
          panel.grid.minor = element_blank(),
          axis.ticks = element_blank(),
          axis.text.x = element_blank(),
          axis.text.y = element_blank(),
          plot.margin = unit(c(-2.5, -1, -1, -1), 'lines'),
          aspect.ratio = 1.5) +
    xlim(c(-1.5, 1.5)) +
    ylim(c(0, 4.25))
}
full_loc_vl <- ggpubr::ggarrange(plotlist = full_pitch_plots_vl, nrow = 1)  


# title style
title_style <- openxlsx::createStyle(halign = 'center', valign = 'center', textDecoration = 'bold', fontSize = 36)
# subtitle style
subtitle_style <- openxlsx::createStyle(halign = 'center', valign = 'center', textDecoration = 'bold', fontSize = 18, fontColour = '#2774AE')
# section header style
subheader_style <- openxlsx::createStyle(halign = 'center', valign = 'center', textDecoration = c('bold', 'italic'), fontSize = 16,
                                         fgFill = '#2774AE', fontColour = 'white', border = c('top', 'bottom', 'left', 'right'), borderStyle = 'medium', borderColour = '#FFD100')
# data header style
header_style <- openxlsx::createStyle(halign = 'center', valign = 'center', textDecoration = 'bold',
                                      border = c('top', 'bottom', 'left', 'right'), borderStyle = 'medium', fontSize = 16)
# data style
data_style <- openxlsx::createStyle(halign = 'center', valign = 'center', textDecoration = 'bold', fontSize = 14,
                                    border = c('top', 'bottom', 'left', 'right'), borderStyle = 'medium')
# notes boxes style
notes_style <- openxlsx::createStyle(halign = 'center', valign = 'center', textDecoration = 'bold', fontSize = 16, fontColour = '#2774AE', wrapText = TRUE,
                                     border = c('top', 'bottom', 'left', 'right'), borderStyle = 'medium')
# loc plot title style
locheader_style <- openxlsx::createStyle(halign = 'center', valign = 'center', textRotation = 90, textDecoration = 'bold', fontSize = 16,
                                         fgFill = '#2774AE', fontColour = 'white', border = c('top', 'bottom', 'left', 'right'), borderStyle = 'medium', borderColour = '#FFD100')

# creating workbook
#wb <- openxlsx::createWorkbook()
# create vs LHB sheet
addWorksheet(wb, paste(name_formatted, 'vs LHB'))

## some sheet formatting
pageSetup(wb, 2, orientation = 'portrait', left = 0.5, right = 0.5, top = 0.5, bottom = 0.5,
          fitToWidth = TRUE, fitToHeight = TRUE, paperSize = 1)

# merging cells
mergeCells(wb, 2, rows = 1, cols = 2:8)
mergeCells(wb, 2, rows = 2, cols = 1:9)
mergeCells(wb, 2, rows = 3, cols = 4:9)
mergeCells(wb, 2, rows = 4, cols = 8:9)
mergeCells(wb, 2, rows = 5, cols = 8:9)
mergeCells(wb, 2, rows = 6, cols = 8:9)
mergeCells(wb, 2, rows = 7, cols = 8:9)
mergeCells(wb, 2, rows = 8, cols = 8:9)
mergeCells(wb, 2, rows = 9, cols = 8:9)
mergeCells(wb, 2, rows = 10, cols = 1:9)
mergeCells(wb, 2, rows = 11, cols = 4:9)
mergeCells(wb, 2, rows = 12:20, cols = 1)
mergeCells(wb, 2, rows = 21:29, cols = 1)
mergeCells(wb, 2, rows = 30:38, cols = 1)
mergeCells(wb, 2, rows = 39:47, cols = 1)
mergeCells(wb, 2, rows = 12:22, cols = 3)
mergeCells(wb, 2, rows = 23:33, cols = 3)
mergeCells(wb, 2, rows = 34:44, cols = 3)
mergeCells(wb, 2, rows = 45:55, cols = 3)

# removing gridlines
# showGridLines(wb, 2, showGridLines = FALSE)

# setting row heights
setRowHeights(wb, 2, rows = 1, heights = 36)
setRowHeights(wb, 2, rows = 2, heights = 30)
setRowHeights(wb, 2, rows = 3, heights = 18)
setRowHeights(wb, 2, rows = 4, heights = 18)
setRowHeights(wb, 2, rows = 5:9, heights = 36)
setRowHeights(wb, 2, rows = 10, heights = 14)
setRowHeights(wb, 2, rows = 11, heights = 18)

# setting column widths
setColWidths(wb, 2, cols = 1, widths = 25)
setColWidths(wb, 2, cols = 2, widths = 2)
setColWidths(wb, 2, cols = 3, widths = 3)
setColWidths(wb, 2, cols = 4, widths = 15)
setColWidths(wb, 2, cols = 5:6, widths = 8)
setColWidths(wb, 2, cols = 7:9, widths = 18)

# inserting TM logo
insertImage(wb, 2, '~/Downloads/TM_logo.png', height = 0.42, width = 2.05, units = 'in', startRow = 1, startCol = 1)

# inserting opp logo
insertImage(wb, 2, '~/Downloads/ucla_opp_logo.png', height = 0.75, width = 1.5, units = 'in', startRow = 1, startCol = 9)

# inserting UCLA (team) logo
insertImage(wb, 2, '~/Downloads/ucla_logo.png', height = 2, width = 2.1, units = 'in', startRow = 49, startCol = 1)

# adding title to sheet
writeData(wb, 2, x = paste(name_formatted, 'vs LHB'), startRow = 1, startCol = 2)
addStyle(wb, 2, rows = 1, cols = 2, style = title_style)

# adding subtitle to sheet
addStyle(wb, 2, rows = 2, cols = 1, style = subtitle_style)

# inserting release plot
writeData(wb, 2, x = 'RELEASE', startRow = 3, startCol = 1)
addStyle(wb, 2, rows = 3, cols = 1, style = subheader_style)
print(vr_plot)
insertPlot(wb, 2, startRow = 4, height = 2.65, width = 2.1, units = 'in')

# inserting pitch table
writeData(wb, 2, x = 'ARSENAL', startRow = 3, startCol = 4)
addStyle(wb, 2, rows = 3, cols = 4:9, style = subheader_style, gridExpand = TRUE)
writeData(wb, 2, pitcher_vl, startRow = 4, startCol = 4)
addStyle(wb, 2, rows = 4, cols = 4:9, style = header_style)
addStyle(wb, 2, rows = which(grepl(pitcher_vl$PITCH, pattern = 'Fastball')) + 4, cols = 4:9,
         style = createStyle(fontColour = '#d22d49', fontSize = 16, textDecoration = 'bold', valign = 'center', halign = 'center',
                             border = c('top', 'bottom', 'left', 'right'), borderStyle = 'thin'))
addStyle(wb, 2, rows = which(grepl(pitcher_vl$PITCH, pattern = 'Sinker')) + 4, cols = 4:9,
         style = createStyle(fontColour = '#fe9d00', fontSize = 16, textDecoration = 'bold', valign = 'center', halign = 'center',
                             border = c('top', 'bottom', 'left', 'right'), borderStyle = 'thin'))
addStyle(wb, 2, rows = which(grepl(pitcher_vl$PITCH, pattern = 'Cutter')) + 4, cols = 4:9,
         style = createStyle(fontColour = '#933f2c', fontSize = 16, textDecoration = 'bold', valign = 'center', halign = 'center',
                             border = c('top', 'bottom', 'left', 'right'), borderStyle = 'thin'))
addStyle(wb, 2, rows = which(grepl(pitcher_vl$PITCH, pattern = 'Curveball')) + 4, cols = 4:9,
         style = createStyle(fontColour = '#00d1ed', fontSize = 16, textDecoration = 'bold', valign = 'center', halign = 'center',
                             border = c('top', 'bottom', 'left', 'right'), borderStyle = 'thin'))
addStyle(wb, 2, rows = which(grepl(pitcher_vl$PITCH, pattern = 'Slider')) + 4, cols = 4:9,
         style = createStyle(fontColour = '#c3bd0d', fontSize = 16, textDecoration = 'bold', valign = 'center', halign = 'center',
                             border = c('top', 'bottom', 'left', 'right'), borderStyle = 'thin'))
addStyle(wb, 2, rows = which(grepl(pitcher_vl$PITCH, pattern = 'ChangeUp')) + 4, cols = 4:9,
         style = createStyle(fontColour = '#23be41', fontSize = 16, textDecoration = 'bold', valign = 'center', halign = 'center',
                             border = c('top', 'bottom', 'left', 'right'), borderStyle = 'thin'))
addStyle(wb, 2, rows = which(grepl(pitcher_vl$PITCH, pattern = 'Splitter')) + 4, cols = 4:9,
         style = createStyle(fontColour = '#3bacac', fontSize = 16, textDecoration = 'bold', valign = 'center', halign = 'center',
                             border = c('top', 'bottom', 'left', 'right'), borderStyle = 'thin'))
addStyle(wb, 2, rows = which(grepl(pitcher_vl$PITCH, pattern = 'Other')) + 4, cols = 4:9,
         style = createStyle(fontColour = '#Acafaf', fontSize = 16, textDecoration = 'bold', valign = 'center', halign = 'center',
                             border = c('top', 'bottom', 'left', 'right'), borderStyle = 'thin'))

# notes section vs LHB
writeData(wb, 2, x = 'NOTES', startRow = 11, startCol = 1)
addStyle(wb, 2, rows = 11, cols = 1, style = subheader_style)
# writeData(wb, 2, 'test 1', startRow = 11, startCol = 1)
addStyle(wb, 2, rows = 12:20, cols = 1, style = notes_style)
# writeData(wb, 2, 'test 2', startRow = 20, startCol = 1)
addStyle(wb, 2, rows = 21:29, cols = 1, style = notes_style)
# writeData(wb, 2, 'test 3', startRow = 29, startCol = 1)
addStyle(wb, 2, rows = 30:38, cols = 1, style = notes_style)
# writeData(wb, 2, 'test 4', startRow = 38, startCol = 1)
addStyle(wb, 2, rows = 39:47, cols = 1, style = notes_style)

# location plots vs LHB
writeData(wb, 2, 'LOCATION', startRow = 11, startCol = 4)
addStyle(wb, 2, rows = 11, cols = 4:9, style = subheader_style)
writeData(wb, 2, 'EARLY', startRow = 12, startCol = 3)
addStyle(wb, 2, rows = 12:22, cols = 3, style = locheader_style)
print(early_loc_vl)
insertPlot(wb, 2, startRow = 12, startCol = 4, height = 2.15, width = 7.35, units = 'in')
writeData(wb, 2, 'HITTERS', startRow = 23, startCol = 3)
addStyle(wb, 2, rows = 23:33, cols = 3, style = locheader_style)
print(hitter_loc_vl)
insertPlot(wb, 2, startRow = 23, startCol = 4, height = 2.15, width = 7.35, units = 'in')
writeData(wb, 2, '2-STRIKE', startRow = 34, startCol = 3)
addStyle(wb, 2, rows = 34:44, cols = 3, style = locheader_style)
print(finish_loc_vl)
insertPlot(wb, 2, startRow = 34, startCol = 4, height = 2.15, width = 7.35, units = 'in')
writeData(wb, 2, 'FULL', startRow = 45, startCol = 3)
addStyle(wb, 2, rows = 45:55, cols = 3, style = locheader_style)
print(full_loc_vl)
insertPlot(wb, 2, startRow = 45, startCol = 4, height = 2.15, width = 7.35, units = 'in')

file_name <- paste0('~/YOUR PATH HERE', name_formatted, '.xlsx')
saveWorkbook(wb, file = file_name, overwrite = TRUE)




###### Genotype × environment interaction and Stability analysis ##############

######################Stability analysis with metan ############################
############################# GGE Model ########################################

library(metan)
library(ggplot2)
library(ggrepel)
library(writexl)
library(openxlsx)
options(max.print = 10000)

############################### data import ####################################

stabdata<-read.csv(file.choose(),)
attach(stabdata)
str(stabdata)
options(max.print = 10000)

######################### factors with unique levels ###########################

stabdata$ENV <- factor(stabdata$ENV, levels=unique(stabdata$ENV))
stabdata$GEN <- factor(stabdata$GEN, levels=unique(stabdata$GEN))
stabdata$REP <- factor(stabdata$REP, levels=unique(stabdata$REP))
str(stabdata)


###################### extract trait name from data file #######################

traitall <- colnames(stabdata)[sapply(stabdata, is.numeric)]
traitall

################### Data inspection and cleaning functions #####################

inspect(stabdata, threshold= 50, plot=FALSE) %>% rmarkdown::paged_table()


for (trait in traitall) {
  find_outliers(stabdata, var = all_of(trait), plots = TRUE)
}

remove_rows_na(stabdata)
replace_zero(stabdata)
find_text_in_num(stabdata, var = all_of(trait))

############################# data analysis ####################################
########################### descriptive stats ##################################

if (!file.exists("output")) {
  dir.create("output")
}

ds <- desc_stat(stabdata, stats="all", hist = TRUE, plot_theme = theme_metan())


write_xlsx(ds, file.path("output", "Descriptive.xlsx"))


############################# mean performances ################################
############################# mean of genotypes ################################

mg <- mean_by(stabdata, GEN) 
mg
View(mg)

############################# mean of environments #############################

me <- mean_by(stabdata, ENV)
me
View(me)

################################# two way mean #################################

dm <- mean_by(stabdata, GEN, ENV)
dm
View(dm)

############### mean performance of genotypes across environments ##############

mge <- stabdata %>% 
  group_by(ENV, GEN) %>%
  desc_stat(stats="mean")
mge
View(mge)

############### Exporting all mean performances computed above #################

write_xlsx(
  list(
    "Genmean" = mg,
    "Envmean" = me,
    "Genmeaninenv" = dm,
    "Genmeaninenv2" = mge
  ),
  file.path("output", "Mean Performance.xlsx")
)

############################ two-way table for all #############################

twgy_list <- list()

for (trait in traitall) {
  twgy_list[[trait]] <- make_mat(stabdata, GEN, ENV, val = trait)
}

result_twgy <- list()
for (trait in traitall) {
  twgy_result<- as.data.frame(twgy_list[[trait]])
  result_twgy[[trait]] <- twgy_result
}

twgy_wb <- createWorkbook()
for (trait in traitall) {
  addWorksheet(twgy_wb, sheetName = paste0(trait, ""))
  writeData(twgy_wb, sheet = trait, x = result_twgy[[trait]], startCol = 2, startRow = 1)
  writeData(twgy_wb, sheet = trait, x = rownames(twgy_list[[trait]]), startCol = 1, startRow = 2)
  writeData(twgy_wb, sheet = trait, x = c("GEN"), startCol = 1, startRow = 1)
}

saveWorkbook(twgy_wb, file.path("output", "TWmean.xlsx"), overwrite = TRUE)

################### plotting performance across environments ###################
#################### make performance for all traits in one ####################
################################## Heatmap #####################################

perfor_heat_list <- list()

for (trait in traitall) {
  perfor_heat_list[[trait]] <-
    ge_plot(
      stabdata,
      ENV,
      GEN,
      !!sym(trait),
      type = 1,
      values = FALSE,
      average = FALSE,
      text_col_pos = c("bottom"),
      text_row_pos = c("left"),
      width_bar = 1.5,
      heigth_bar = 20,
      xlab = "ENV",
      ylab = "GEN",
      plot_theme = theme_metan(),
      colour = TRUE
    ) + geom_tile(color = "transparent") + labs(title = paste0(trait, " performance across eight environments")) + theme(legend.title = element_text(), axis.text.x.bottom = element_text(angle = 0, hjust = .5)) + guides(fill = guide_colourbar(title = trait, barwidth = 1.5, barheight = 20))
  assign(paste0(trait, "_perfor_heat"), perfor_heat_list[[trait]])
}

############################## print all plots once ############################

for (trait in traitall) {
  print(perfor_heat_list[[trait]])
}

##################### high quality save all in one  ############################

if (!file.exists("output")) {
  dir.create("output")
}

if (!file.exists(file.path("output", "perfor_heat_plot"))) {
  dir.create(file.path("output", "perfor_heat_plot"))
}

for (trait in traitall) {
  ggsave(filename = file.path("output", "perfor_heat_plot", paste0(trait, ".png")),
         plot = perfor_heat_list[[trait]], width = 20, height = 30,
         dpi = 600, units = "cm")
}

################################# Line plot ####################################

perfor_line_list <- list()

for (trait in traitall) {
  perfor_line_list[[trait]] <-
    ge_plot(
      stabdata,
      ENV,
      GEN,
      !!sym(trait),
      type = 1,
      values = FALSE,
      average = FALSE,
      text_col_pos = c("bottom"),
      text_row_pos = c("left"),
      width_bar = 1.5,
      heigth_bar = 20,
      xlab = "ENV",
      ylab = "GEN",
      plot_theme = theme_metan(),
      colour = TRUE
    ) + geom_tile(color = "transparent") + labs(title = paste0(trait, " performance across eight environments")) + theme(legend.title = element_text(), axis.text.x.bottom = element_text(angle = 0, hjust = .5)) + guides(fill = guide_colourbar(title = trait, barwidth = 1.5, barheight = 20))
  assign(paste0(trait, "_perfor_line"), perfor_line_list[[trait]])
}

############################## print all plots once ############################

for (trait in traitall) {
  print(perfor_line_list[[trait]])
}

##################### high quality save all in one  ############################

if (!file.exists("output")) {
  dir.create("output")
}

if (!file.exists(file.path("output", "perfor_line_plot"))) {
  dir.create(file.path("output", "perfor_line_plot"))
}

for (trait in traitall) {
  ggsave(filename = file.path("output", "perfor_line_plot", paste0(trait, ".png")),
         plot = perfor_line_list[[trait]], width = 20, height = 30,
         dpi = 600, units = "cm")
}

########################## Genotype-environment winners ########################
#edit better argument as for some variables lower values are preferred and higher for others ####

traitall ### view your traits to decide for above condition ###

win <-
  ge_winners(
    stabdata,
    ENV,
    GEN,
    resp = everything(),
    better = c(h, h, h, h, h, h, h, h, l, l, l, h, h, h, h, h, h, h, l) 
  )
win
View(win)
ranks <-
  ge_winners(
    stabdata,
    ENV,
    GEN,
    resp = everything(),
    type = "ranks",
    better = c(h, h, h, h, h, h, h, h, l, l, l, h, h, h, h, h, h, h, l)
  )
ranks
View(ranks)

write_xlsx(list("winner" = win, "ranks" = ranks), file.path("output", "winner rank.xlsx"))

############################ ge or gge effects #################################
######################### combined for all ge effects ##########################
########################## ge effects to excel #################################

ge_list <- ge_effects(stabdata, ENV, GEN, resp = everything(), type = "ge")

result_ge_list <- list()
for (trait in traitall) {
  ge_list_result <- as.data.frame(ge_list[[trait]])
  result_ge_list[[trait]] <- ge_list_result
}

########################### save all in one excel  #############################

ge_list_wb <- createWorkbook()
for (trait in traitall) {
  addWorksheet(ge_list_wb, sheetName = paste0(trait, ""))
  writeData(ge_list_wb, sheet = trait, x = result_ge_list[[trait]])
}
saveWorkbook(ge_list_wb, file.path("output","ge_effects.xlsx"), overwrite = TRUE)

############################## ge effects plots ################################

ge_plots <- list()

for (trait in traitall) {
  ge_plots[[trait]] <- plot(ge_list) + aes(ENV, GEN) + theme(legend.title = element_text()) + guides(fill = guide_colourbar(title = paste0(trait, " ge effects"), barwidth = 1.5, barheight = 20))
}  ## also coord_flip() in place of aes

################# print all plots once #########################################

for (trait in traitall) {
  print(ge_plots[[trait]])
}

##################### high quality save all in one  ############################

if (!file.exists("output")) {
  dir.create("output")
}

if (!file.exists(file.path("output", "ge_plots"))) {
  dir.create(file.path("output", "ge_plots"))
}

for (trait in traitall) {
  ggsave(filename = file.path("output", "ge_plots", paste0(trait, ".png")),
         plot = ge_plots[[trait]], width = 20, height = 30,
         dpi = 600, units = "cm")
}

########################## gge effects to excel ################################

gge_list <- ge_effects(stabdata, ENV, GEN, resp = everything(), type = "gge")

result_gge_list <- list()
for (trait in traitall) {
  gge_list_result <- as.data.frame(gge_list[[trait]])
  result_gge_list[[trait]] <- gge_list_result
}

########################### save all in one excel  #############################

gge_list_wb <- createWorkbook()
for (trait in traitall) {
  addWorksheet(gge_list_wb, sheetName = paste0(trait, ""))
  writeData(gge_list_wb, sheet = trait, x = result_gge_list[[trait]])
}
saveWorkbook(gge_list_wb, file.path("output","gge_effects.xlsx"), overwrite = TRUE)

############################## ge effects plots ################################

gge_plots <- list()

for (trait in traitall) {
  gge_plots[[trait]] <- plot(gge_list) + aes(ENV, GEN) + theme(legend.title = element_text()) + guides(fill = guide_colourbar(title = paste0(trait, " gge effects"), barwidth = 1.5, barheight = 20))
}  ## also coord_flip() in place of aes

################# print all plots once #########################################

for (trait in traitall) {
  print(gge_plots[[trait]])
}

##################### high quality save all in one  ############################

if (!file.exists("output")) {
  dir.create("output")
}

if (!file.exists(file.path("output", "gge_plots"))) {
  dir.create(file.path("output", "gge_plots"))
}

for (trait in traitall) {
  ggsave(filename = file.path("output", "gge_plots", paste0(trait, ".png")),
         plot = gge_plots[[trait]], width = 20, height = 30,
         dpi = 600, units = "cm")
}

############################ fixed effect models ###############################
########################### Individual  and Joint anova ########################
######################### Individual anova for all traits ######################

aovind_list <- anova_ind(stabdata, env = ENV, gen = GEN, rep = REP, resp = everything())

result_aovind <- list()
for (trait in traitall) {
  ind_result<- as.data.frame(aovind_list[[trait]]$individual)
  result_aovind[[trait]] <- ind_result
}

aovind_wb <- createWorkbook()
for (trait in traitall) {
  addWorksheet(aovind_wb, sheetName = paste0(trait, ""))
  writeData(aovind_wb, sheet = trait, x = result_aovind[[trait]])
}
saveWorkbook(aovind_wb, file.path("output","indaovall.xlsx"), overwrite = TRUE)

################## Joint anova for all traits (ANOVA) ##########################

aovjoin_list <- anova_joint(stabdata, env = ENV, gen = GEN, rep = REP, resp = everything())

result_aovjoin <- list()
for (trait in traitall) {
  join_result<- as.data.frame(aovjoin_list[[trait]]$anova)
  result_aovjoin[[trait]] <- join_result
}

aovjoin_wb <- createWorkbook()
for (trait in traitall) {
  addWorksheet(aovjoin_wb, sheetName = paste0(trait, ""))
  writeData(aovjoin_wb, sheet = trait, x = result_aovjoin[[trait]])
}
saveWorkbook(aovjoin_wb, file.path("output","joinaovall.xlsx"), overwrite = TRUE)

################## Joint anova for all traits (Details) (2) ####################

result_aovjoin2 <- list()
for (trait in traitall) {
  join_result2<- as.data.frame(aovjoin_list[[trait]]$details)
  result_aovjoin2[[trait]] <- join_result2
}

aovjoin_wb2 <- createWorkbook()
for (trait in traitall) {
  addWorksheet(aovjoin_wb2, sheetName = paste0(trait, ""))
  writeData(aovjoin_wb2, sheet = trait, x = result_aovjoin2[[trait]])
}
saveWorkbook(aovjoin_wb2, file.path("output", "joinaovall2.xlsx"), overwrite = TRUE)

#################### GGE based stability analysis for all ######################
##################### (Centering is fixed to Environment) ######################
#### below three SVP set scenarios are estimated to plot various gge biplots####

############################## SVP set to genotype #############################
######## Scalling = 0, centering = 2 (Environment), SVP = 1 (Genotype) #########

gge_gen_list <-
  gge(
    stabdata,
    ENV,
    GEN,
    resp = everything(),
    scaling = "none",
    centering = "environment", ## "global" = 1, "environment" = 2, "double = 3"
    svp = "genotype"           ## "genotype" = 1,
  )

############################ SVP set to environment ############################
###### Scalling = 0, centering = 2 (Environment), SVP = 2 (Environment) ########

gge_env_list <-
  gge(
    stabdata,
    ENV,
    GEN,
    resp = everything(),
    scaling = "none",
    centering = "environment", ## "global" = 1, "environment" = 2, "double = 3"
    svp = "environment"        ## "environment" = 2, 
  )

############################# SVP set to symmetrical ###########################
###### Scalling = 0, centering = 2 (Environment), SVP = 2 (symmetrical) ########

gge_sym_list <-
  gge(
    stabdata,
    ENV,
    GEN,
    resp = everything(),
    scaling = "none",
    centering = "environment", ## "global" = 1, "environment" = 2, "double = 3"
    svp = "symmetrical"        ## "symmetrical = 3"
  )

############# GGE analysis produce 10 different type of biplots ################
###### help https://tiagoolivoto.github.io/metan/reference/plot.gge.html #######
##### help https://tiagoolivoto.github.io/metan/articles/vignettes_gge.html ####


######################### function to generate GGE biplot ######################

create_gge_plots <- function(gge_list, traitall, plot_type, sel_env = NULL, sel_gen = NULL, sel_gen1 = NULL, sel_gen2 = NULL, title_text = NULL,...) {
  # Ensure plot_types is always a vector
  if (!is.vector(plot_type)) {
    plot_type <- as.vector(plot_type)
  }
  
  gge_plot_lists <- list()
  for (i in seq_along(plot_type)) {
    gge_plot_list <- list()
    for (trait in traitall) {
      if (plot_type[i] %in% c(5, 7, 9)) {
        gge_plot_list[[trait]] <- plot(gge_list,
                                       var = trait,
                                       type = plot_type[i],
                                       repel = TRUE,
                                       repulsion = 1,
                                       max_overlaps = 50,
                                       sel_env = ifelse(plot_type[i] == 5, sel_env, NA),
                                       sel_gen = ifelse(plot_type[i] == 7, sel_gen, NA),
                                       sel_gen1 = ifelse(plot_type[i] == 9, sel_gen1, NA),
                                       sel_gen2 = ifelse(plot_type[i] == 9, sel_gen2, NA),
                                       shape.gen = 21,
                                       shape.env = 23,
                                       line.type.gen = "dotted",
                                       size.shape = 2.2,
                                       size.shape.win = 3.4,
                                       size.stroke = 0.2,
                                       col.stroke = "white",
                                       col.gen = "#215C29",
                                       col.env = "#F68A31",
                                       col.line = "#F68A31",
                                       col.alpha = 1,
                                       col.circle = "gray",
                                       col.alpha.circle = 0.5,
                                       size.text.gen = if (plot_type[i] == 7) 4.5 else 2,
                                       size.text.env = if (plot_type[i] == 5) 4.5 else 2,
                                       size.text.lab = 12,
                                       size.text.win = 4.5,
                                       size.line = 0.5,
                                       axis_expand = 2,
                                       title = TRUE,
                                       plot_theme = theme_metan(),
                                       leg.lab = c("Environment", "Genotype")) + labs(title = paste0(title_text, trait)) + theme(
                                         plot.title = element_text(color = "black"),
                                         panel.background = element_rect(fill = "white"), plot.background = element_rect(fill = "white"),
                                         legend.position = c(0.88, .05),
                                         legend.background = element_rect(fill = NA))
      } else {
        # Handle other plot types here, using additional arguments passed to the plot function
        gge_plot_list[[trait]] <- plot(gge_list,
                                       var = trait,
                                       type = plot_type[i],
                                       repel = TRUE,
                                       repulsion = 1,
                                       max_overlaps = 50,
                                       sel_env = sel_env,
                                       sel_gen = sel_gen,
                                       sel_gen1 = sel_gen1,
                                       sel_gen2 = sel_gen2,
                                       shape.gen = 21,
                                       shape.env = 23,
                                       line.type.gen = "dotted",
                                       size.shape = 2.2,
                                       size.shape.win = 3.4,
                                       size.stroke = 0.2,
                                       col.stroke = "white",
                                       col.gen = "#215C29",
                                       col.env = "#F68A31",
                                       col.line = "#F68A31",
                                       col.alpha = 1,
                                       col.circle = "gray",
                                       col.alpha.circle = 0.5,
                                       size.text.gen = 3.5,
                                       size.text.env = 3.5,
                                       size.text.lab = 12,
                                       size.text.win = 4.5,
                                       size.line = 0.5,
                                       axis_expand = 1.2,
                                       title = TRUE,
                                       plot_theme = theme_metan(),
                                       leg.lab = c("Environment", "Genotype"),
                                       ...) + labs(title = paste0(title_text , trait)) + theme(
                                         plot.title = element_text(color = "black"),
                                         panel.background = element_rect(fill = "white"), plot.background = element_rect(fill = "white"),
                                         legend.position = c(0.88, .05),
                                         legend.background = element_rect(fill = NA))
      }
    }
    # Generate a unique name for each output list
    list_name <- paste0("gge_", plot_type[i], "_list")
    gge_plot_lists[[list_name]] <- gge_plot_list
    assign(list_name, gge_plot_list, envir = .GlobalEnv)
  }
  return(gge_plot_lists)
}

######################### function to save GGE biplot ##########################

save_gge_plots <- function(gge_save_list, traitall, folder) {
  output_folder <- file.path("output", folder)
  if (!dir.exists("output")) {
    dir.create("output")
  }
  if (!dir.exists(output_folder)) {
    dir.create(output_folder)
  }
  for (trait in traitall) {
    ggsave(filename = file.path(output_folder, paste0(trait, ".png")),
           plot = gge_save_list[[trait]], width = 15, height = 15,
           dpi = 600, units = "cm")
  }
}


####################### GGE biplots for all in one #############################

############################ (type = 1 A basic biplot) #########################

create_gge_plots(
  gge_list = gge_env_list,
  traitall = traitall,
  plot_type = 1,
  sel_env = NULL,
  sel_gen = NULL,
  sel_gen1 = NA,
  sel_gen2 = NA,
  title_text = "GGE Biplot for "
)

############################## print all plots once ############################

for (trait in traitall) {
  print(gge_1_list[[trait]])
}

##################### high quality save all in one  ############################

save_gge_plots(gge_save_list = gge_1_list, traitall = traitall, folder = "gge_1_plots")

############### (type = 2 Mean performance vs. stability.) #####################

create_gge_plots(
  gge_list = gge_gen_list,
  traitall = traitall,
  plot_type = 2,
  sel_env = NULL,
  sel_gen = NULL,
  sel_gen1 = NA,
  sel_gen2 = NA,
  title_text = "Mean vs Stability for "
)

############################## print all plots once ############################

for (trait in traitall) {
  print(gge_2_list[[trait]])
}

##################### high quality save all in one  ############################

save_gge_plots(gge_save_list = gge_2_list, traitall = traitall, folder = "gge_2_plots")

########################### (type = 3 Which-won-where) #########################

create_gge_plots(
  gge_list = gge_sym_list,
  traitall = traitall,
  plot_type = 3,
  sel_env = NULL,
  sel_gen = NULL,
  sel_gen1 = NA,
  sel_gen2 = NA,
  title_text = "Which-Won-Where View of the GGE Biplot for "
)

############################## print all plots once ############################

for (trait in traitall) {
  print(gge_3_list[[trait]])
}

##################### high quality save all in one  ############################

save_gge_plots(gge_save_list = gge_3_list, traitall = traitall, folder = "gge_3_plots")

########## (type = 4 Discriminativeness vs. representativeness) ################

create_gge_plots(
  gge_list = gge_sym_list,
  traitall = traitall,
  plot_type = 4,
  sel_env = NULL,
  sel_gen = NULL,
  sel_gen1 = NA,
  sel_gen2 = NA,
  title_text = "Discriminativeness vs Representativeness for "
)

############################## print all plots once ############################

for (trait in traitall) {
  print(gge_4_list[[trait]])
}

##################### high quality save all in one  ############################

save_gge_plots(gge_save_list = gge_4_list, traitall = traitall, folder = "gge_4_plots")

######################## (type = 5 Examine an environment) #####################

create_gge_plots(
  gge_list = gge_sym_list,
  traitall = traitall,
  plot_type = 5,
  sel_env = "HTS21", #### select environment to examine, must match env name, insert in quotation marks
  sel_gen = NULL,
  sel_gen1 = NA,
  sel_gen2 = NA,
  title_text = "Examining Environment: HTS21 for " ### insert name of environment in place of X
)

############################## print all plots once ############################

for (trait in traitall) {
  print(gge_5_list[[trait]])
}

##################### high quality save all in one  ############################

save_gge_plots(gge_save_list = gge_5_list, traitall = traitall, folder = "gge_5_plots")

######################## (type = 6 Ranking environments) #######################

create_gge_plots(
  gge_list = gge_env_list,
  traitall = traitall,
  plot_type = 6,
  sel_env = NULL,
  sel_gen = NULL,
  sel_gen1 = NA,
  sel_gen2 = NA,
  title_text = "Ranking Environments for "
)

############################## print all plots once ############################

for (trait in traitall) {
  print(gge_6_list[[trait]])
}

##################### high quality save all in one  ############################

save_gge_plots(gge_save_list = gge_6_list, traitall = traitall, folder = "gge_6_plots")

######################### (type = 7 Examine a genotype) ########################

create_gge_plots(
  gge_list = gge_gen_list,
  traitall = traitall,
  plot_type = 7,
  sel_env = NULL,
  sel_gen = "C306", #### select genotype to examine, must match gen name, insert in quotation marks
  sel_gen1 = NA,
  sel_gen2 = NA,
  title_text = "Examining Genotype: C306 for " ### insert name of genotype in place of X
)

############################## print all plots once ############################

for (trait in traitall) {
  print(gge_7_list[[trait]])
}

##################### high quality save all in one  ############################

save_gge_plots(gge_save_list = gge_7_list, traitall = traitall, folder = "gge_7_plots")

######################### (type = 8 Ranking genotypes) #########################

create_gge_plots(
  gge_list = gge_gen_list,
  traitall = traitall,
  plot_type = 8,
  sel_env = NULL,
  sel_gen = NULL,
  sel_gen1 = NA,
  sel_gen2 = NA,
  title_text = "Ranking Genotypes for "
)

############################## print all plots once ############################

for (trait in traitall) {
  print(gge_8_list[[trait]])
}

##################### high quality save all in one  ############################

save_gge_plots(gge_save_list = gge_8_list, traitall = traitall, folder = "gge_8_plots")

####################### (type = 9 Compare two genotypes) #######################

create_gge_plots(
  gge_list = gge_sym_list,
  traitall = traitall,
  plot_type = 9,
  sel_env = NULL,
  sel_gen = NULL,
  sel_gen1 = "DBW222", #### select 1st genotype to compare
  sel_gen2 = "DBW187", #### select 2nd genotype to compare
  title_text = "Comparison Among Genotype DBW222 : Genotpe 187 for " ## change X1 and X2 with genotype name  
)

############################## print all plots once ############################

for (trait in traitall) {
  print(gge_9_list[[trait]])
}

##################### high quality save all in one  ############################

save_gge_plots(gge_save_list = gge_9_list, traitall = traitall, folder = "gge_9_plots")

############### (type = 10 Relationship among environments) ####################

create_gge_plots(
  gge_list = gge_env_list,
  traitall = traitall,
  plot_type = 10,
  sel_env = NULL,
  sel_gen = NULL,
  sel_gen1 = NA, 
  sel_gen2 = NA, 
  title_text = "Relationship Among Environments for "
)

############################## print all plots once ############################

for (trait in traitall) {
  print(gge_10_list[[trait]])
}

##################### high quality save all in one  ############################

save_gge_plots(gge_save_list = gge_10_list, traitall = traitall, folder = "gge_10_plots")

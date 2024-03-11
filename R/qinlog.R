#' @title 单+多因素logit回归
#' @description
#' 该函数可用于进行批量单因素logit回归分析以及多因素logit回归分析。
#' @param data *输入数据框
#' @param vars *自变量
#' @param y *结局变量
#' @param reg 回归方式："uni","mul","uni+mul"（默认）
#' @param decimal 结果保留小数点位，默认为3
#' @param P_decimal P值保留小数点位，默认为3
#' @param sig_P 多因素回归纳入变量P值设定，默认为0.05
#' @param CL 置信度，默认为0.95
#' @param direction 逐步回归方向："both", "backward", "forward",参数缺失则不进行逐步回归
#' @author qinstudent
#' @export
#' @importFrom dplyr %>% mutate across select add_row na_if
#' @importFrom tidyr unnest
#' @importFrom purrr map
#' @importFrom tibble rownames_to_column
#' @importFrom stringr str_remove str_split str_locate str_c str_sub str_length
#' @importFrom officer block_section read_docx body_add_par body_end_block_section
#' @importFrom flextable flextable color padding add_footer_lines set_table_properties body_add_flextable set_flextable_defaults
qinlog <- function(data=NULL,vars=NULL,y=NULL,reg = NULL,decimal = NULL,P_decimal = NULL,sig_P = NULL,CL = NULL,direction = NULL)
  {
  ####1.预处理####
  #默认值
  if (is.null(data)) stop("Missing data input")
  if (is.null(vars)) stop("Missing vars input")
  if (is.null(y)) stop("Missing y input")
  if (is.null(decimal)) decimal = 3
  if (is.null(P_decimal)) P_decimal = 3
  if (is.null(CL)) CL = 0.95
  if (is.null(reg)) reg = "uni+mul"
  if (!reg %in% c("uni","mul","uni+mul")) stop('The input regression style must be within "uni","mul","uni+mul"')
  if (is.null(sig_P) & (reg %in% c("mul","uni+mul"))) {
    sig_P = 0.05
    print("注：多因素回归变量筛选P值设定为默认值0.05")
  }
  #剔除结局缺失
  if (nrow(subset(data,is.na(data[,y])))>0) {
    data <- subset(data,!is.na(data[,y]))
    print(paste0("回归剔除",nrow(subset(data,is.na(data[,y]))),"例结局缺失"))
  }
  ####2.单因素回归####
  if (reg %in% c("uni","uni+mul")) {
  reg_table <- data.frame(vars) %>%
    mutate(reg_model = map(data[vars],~ glm(as.formula(paste0(y, "~", ".x")),data = data,family = binomial())),
    reg_model_sum = map(reg_model,~ summary(.x) %>% coefficients %>% as.data.frame %>% rownames_to_column),
    exp_b = map(reg_model,~ coef(.x) %>% exp)) %>% #建模
    unnest(c(reg_model_sum,exp_b)) %>% #解压
    mutate(b = sprintf(paste0('%0.',decimal,'f'),Estimate),
           z = sprintf(paste0('%0.',decimal,'f'),`z value`),
           P = sprintf(paste0('%0.',P_decimal,'f'),`Pr(>|z|)`),
           P = ifelse(`Pr(>|z|)` < 10^(-P_decimal),paste0("<",10^(-P_decimal)),P),
           S.E = sprintf(paste0('%0.',decimal,'f'),`Std. Error`),
           exp_b_lower = exp(Estimate - qnorm(1-(1-CL)/2)*`Std. Error`) %>% round(digits = decimal),
           exp_b_lower = ifelse(exp_b_lower<1 & 1-exp_b_lower<10^(-decimal),1-10^(-decimal),
                                ifelse(exp_b_lower>1 & exp_b_lower-1<10^(-decimal),1+10^(-decimal),exp_b_lower)),
           exp_b_lower = sprintf(paste0('%0.',decimal,'f'),exp_b_lower),
           exp_b_upper = exp(Estimate + qnorm(1-(1-CL)/2)*`Std. Error`) %>% round(digits = decimal),
           exp_b_upper = ifelse(exp_b_upper<1 & 1-exp_b_upper<10^(-decimal),1-10^(-decimal),
                                ifelse(exp_b_upper>1 & exp_b_upper-1<10^(-decimal),1+10^(-decimal),exp_b_upper)),
           exp_b_upper = sprintf(paste0('%0.',decimal,'f'),exp_b_upper),
           exp_b = round(exp_b,digits = decimal),
           exp_b = ifelse(exp_b<1 & 1-exp_b<10^(-decimal),1-10^(-decimal),
                          ifelse(exp_b>1 & exp_b-1<10^(-decimal),1+10^(-decimal),exp_b)),
           exp_b = sprintf(paste0('%0.',decimal,'f'),exp_b),
           exp_b_CI = str_c(exp_b, ' (', exp_b_lower, '-',exp_b_upper, ')'),
           Variables = vars,
           across(rowname, ~ str_remove(.x, '.x'))) %>% #取值
    filter(rowname != "(Intercept)")
  #筛出显著变量以备多因素回归
  sig_row <- which(reg_table$`Pr(>|z|)` < sig_P)
  sig_vars <- (reg_table[sig_row,1] %>% unique %>% as.data.frame)[,1]
  #筛出P<0.05变量以备文字描述
  sig_row_description <- which(reg_table$`Pr(>|z|)` < 0.05)
  sig_vars_description <- (reg_table[sig_row_description,1] %>% unique %>% as.data.frame)[,1]
  #调整表格
  reg_table %>% select(Variables,rowname,b,S.E,z,P, exp_b_CI,`Pr(>|z|)`,exp_b_lower,exp_b_upper) %>% as.data.frame -> reg_table
  catrow <- which((reg_table$Variables %in% vars[unlist(unname(do.call(cbind.data.frame, lapply(data[vars],is.factor))))])& (!duplicated(reg_table$Variables)))
  seq <- (0:(length(catrow)-1))*2
  for (i in catrow+seq) {
    reg_table <- reg_table %>%
      add_row(Variables = c(reg_table[i,1],"  "),rowname = c("",levels(data[,reg_table[i,1]])[1]),b=c("",""),S.E=c("",""),z=c("",""),
              P=c("",""),exp_b_CI=c("","Reference"),.before = i)}
  catrow2 <- which(duplicated(reg_table$Variables)&(reg_table$Variables!="  "))
  reg_table[catrow2,"Variables"] <- "  "
  reg_table <- reg_table %>% mutate(Variables = paste0(Variables,rowname)) %>% select(-rowname)
  reg_table$Variables <- gsub("\t","  ",reg_table$Variables)
  names(reg_table)[6] <- paste0("OR(",CL*100,"%CI)")
  }
  ####3.多因素回归####
  #####3.1多因素####
  if ((reg %in% c("mul","uni+mul"))&is.null(direction)) {
    if (reg == "mul") {mul_vars <- vars}else mul_vars <- sig_vars
    mul_reg_model <- glm(as.formula(paste0(y,"~",paste(mul_vars,collapse = "+"))),data = data,family = binomial())
    mul_reg_model_sum <- summary(mul_reg_model) %>% coefficients %>% as.data.frame %>% rownames_to_column %>% mutate(exp_b = exp(Estimate))
    mul_reg_variable <- data.frame(mul_vars) %>%
      mutate(levels = map(data[mul_vars],levels)) %>%
      unnest(levels) %>%
      mutate(rowname = paste0(mul_vars,levels))
    mul_reg_merge <- merge(mul_reg_variable,mul_reg_model_sum,by="rowname",all=T)
    mul_reg_table <- mul_reg_merge[match(mul_reg_variable$rowname, mul_reg_merge$rowname),] %>% #merge数据框按x行序
      mutate(b = sprintf(paste0('%0.',decimal,'f'),Estimate),
             z = sprintf(paste0('%0.',decimal,'f'),`z value`),
             P = sprintf(paste0('%0.',P_decimal,'f'),`Pr(>|z|)`),
             P = ifelse(`Pr(>|z|)` < 10^(-P_decimal),paste0("<",10^(-P_decimal)),P),
             S.E = sprintf(paste0('%0.',decimal,'f'),`Std. Error`),
             exp_b_lower = exp(Estimate - qnorm(1-(1-CL)/2)*`Std. Error`) %>% round(digits = decimal),
             exp_b_lower = ifelse(exp_b_lower<1 & 1-exp_b_lower<10^(-decimal),1-10^(-decimal),
                                  ifelse(exp_b_lower>1 & exp_b_lower-1<10^(-decimal),1+10^(-decimal),exp_b_lower)),
             exp_b_lower = sprintf(paste0('%0.',decimal,'f'),exp_b_lower),
             exp_b_upper = exp(Estimate + qnorm(1-(1-CL)/2)*`Std. Error`) %>% round(digits = decimal),
             exp_b_upper = ifelse(exp_b_upper<1 & 1-exp_b_upper<10^(-decimal),1-10^(-decimal),
                                  ifelse(exp_b_upper>1 & exp_b_upper-1<10^(-decimal),1+10^(-decimal),exp_b_upper)),
             exp_b_upper = sprintf(paste0('%0.',decimal,'f'),exp_b_upper),
             exp_b = round(exp_b,digits = decimal),
             exp_b = ifelse(exp_b<1 & 1-exp_b<10^(-decimal),1-10^(-decimal),
                            ifelse(exp_b>1 & exp_b-1<10^(-decimal),1+10^(-decimal),exp_b)),
             exp_b = sprintf(paste0('%0.',decimal,'f'),exp_b),
             exp_b_CI = str_c(exp_b, ' (', exp_b_lower, '-',exp_b_upper, ')'),
             Variables = mul_vars
             )
    #筛出P<0.05变量以备文字描述
    mul_sig_row_description <- which(mul_reg_table$`Pr(>|z|)` < 0.05)
    mul_sig_vars_description <- (mul_reg_table[mul_sig_row_description,1] %>% unique %>% as.data.frame)[,1]
    #调整表格
    mul_reg_table %>% select(Variables,levels,b,S.E,z,P, exp_b_CI,`Pr(>|z|)`,exp_b_lower,exp_b_upper) %>% as.data.frame -> mul_reg_table
    catrow_mtable <- which((mul_reg_table$Variables %in% vars[unlist(unname(do.call(cbind.data.frame, lapply(data[vars],is.factor))))])& (!duplicated(mul_reg_table$Variables)))
    mul_reg_table[catrow_mtable,"exp_b_CI"] = "Reference"
    seq2 <- 0:(length(catrow_mtable)-1)
    for (i in catrow_mtable+seq2) {
      mul_reg_table <- mul_reg_table %>%
        add_row(Variables = mul_reg_table[i,1],levels = "",b="",S.E="",z="",
                P="",.before = i)}
    catrow_mtable2 <- which(duplicated(mul_reg_table$Variables))
    mul_reg_table[catrow_mtable2,"Variables"] <- "  "
    mul_reg_table <- mul_reg_table %>% mutate(Variables = paste0(Variables,levels)) %>% select(-levels)
    mul_reg_table$Variables <- gsub("\t","  ",mul_reg_table$Variables)
    names(mul_reg_table)[6] <- paste0("OR(",CL*100,"%CI)")
    mul_reg_table$`b` <- na_if(mul_reg_table$`b`, "NA")
    mul_reg_table$`S.E` <- na_if(mul_reg_table$`S.E`, "NA")
    mul_reg_table$`z` <- na_if(mul_reg_table$`z`, "NA")
    mul_reg_table$`exp_b_lower` <- na_if(mul_reg_table$`exp_b_lower`, "NA") %>% as.numeric
    mul_reg_table$`exp_b_upper` <- na_if(mul_reg_table$`exp_b_upper`, "NA") %>% as.numeric
  }
  #####3.2多因素逐步####
  if ((reg %in% c("mul","uni+mul"))&(!is.null(direction))) {
    if (reg == "mul") {mul_vars <- vars}else mul_vars <- sig_vars
    mul_reg_model <- glm(as.formula(paste0(y,"~",paste(mul_vars,collapse = "+"))),data = data,family = binomial())
    step_mul_reg_model_sum <- step(mul_reg_model,direction = direction,trace = FALSE) %>% summary %>% coefficients %>% as.data.frame %>% rownames_to_column %>% mutate(exp_b = exp(Estimate))
    step_mul_reg_model_call <- summary(step(mul_reg_model,direction = direction,trace = FALSE))$call %>% as.character
    step_mul_vars <- str_sub(step_mul_reg_model_call[2],str_locate(step_mul_reg_model_call[2],pattern = "~")[1,1]+2,str_length(step_mul_reg_model_call[2])) %>%
      str_split(pattern = " \\+ ",simplify = T) %>% as.vector
    mul_reg_variable2 <- data.frame(step_mul_vars) %>%
      mutate(levels = map(data[step_mul_vars],levels)) %>%
      unnest(levels) %>%
      mutate(rowname = paste0(step_mul_vars,levels))
    mul_reg_merge2 <- merge(mul_reg_variable2,step_mul_reg_model_sum,by="rowname",all=T)
    mul_reg_table2 <- mul_reg_merge2[match(mul_reg_variable2$rowname, mul_reg_merge2$rowname),] %>% #merge数据框按x行序
      mutate(b = sprintf(paste0('%0.',decimal,'f'),Estimate),
             z = sprintf(paste0('%0.',decimal,'f'),`z value`),
             P = sprintf(paste0('%0.',P_decimal,'f'),`Pr(>|z|)`),
             P = ifelse(`Pr(>|z|)` < 10^(-P_decimal),paste0("<",10^(-P_decimal)),P),
             S.E = sprintf(paste0('%0.',decimal,'f'),`Std. Error`),
             exp_b_lower = exp(Estimate - qnorm(1-(1-CL)/2)*`Std. Error`) %>% round(digits = decimal),
             exp_b_lower = ifelse(exp_b_lower<1 & 1-exp_b_lower<10^(-decimal),1-10^(-decimal),
                                  ifelse(exp_b_lower>1 & exp_b_lower-1<10^(-decimal),1+10^(-decimal),exp_b_lower)),
             exp_b_lower = sprintf(paste0('%0.',decimal,'f'),exp_b_lower),
             exp_b_upper = exp(Estimate + qnorm(1-(1-CL)/2)*`Std. Error`) %>% round(digits = decimal),
             exp_b_upper = ifelse(exp_b_upper<1 & 1-exp_b_upper<10^(-decimal),1-10^(-decimal),
                                  ifelse(exp_b_upper>1 & exp_b_upper-1<10^(-decimal),1+10^(-decimal),exp_b_upper)),
             exp_b_upper = sprintf(paste0('%0.',decimal,'f'),exp_b_upper),
             exp_b = round(exp_b,digits = decimal),
             exp_b = ifelse(exp_b<1 & 1-exp_b<10^(-decimal),1-10^(-decimal),
                            ifelse(exp_b>1 & exp_b-1<10^(-decimal),1+10^(-decimal),exp_b)),
             exp_b = sprintf(paste0('%0.',decimal,'f'),exp_b),
             exp_b_CI = str_c(exp_b, ' (', exp_b_lower, '-',exp_b_upper, ')'),
             Variables = step_mul_vars
      )
    #筛出P<0.05变量以备文字描述
    mul_sig_row_description <- which(mul_reg_table2$`Pr(>|z|)` < 0.05)
    mul_sig_vars_description <- (mul_reg_table2[mul_sig_row_description,1] %>% unique %>% as.data.frame)[,1]
    #调整表格
    mul_reg_table2 %>% select(Variables,levels,b,S.E,z,P, exp_b_CI,`Pr(>|z|)`,exp_b_lower,exp_b_upper) %>% as.data.frame -> mul_reg_table2
    step_catrow_mtable <- which((mul_reg_table2$Variables %in% vars[unlist(unname(do.call(cbind.data.frame, lapply(data[vars],is.factor))))])& (!duplicated(mul_reg_table2$Variables)))
    mul_reg_table2[step_catrow_mtable,"exp_b_CI"] = "Reference"
    seq3 <- 0:(length(step_catrow_mtable)-1)
    for (i in step_catrow_mtable+seq3) {
      mul_reg_table2 <- mul_reg_table2 %>%
        add_row(Variables = mul_reg_table2[i,1],levels = "",.before = i)}
    step_catrow_mtable2 <- which(duplicated(mul_reg_table2$Variables))
    mul_reg_table2[step_catrow_mtable2,"Variables"] <- "  "
    mul_reg_table2 <- mul_reg_table2 %>% mutate(Variables = paste0(Variables,levels)) %>% select(-levels)
    mul_reg_table2$Variables <- gsub("\t","  ",mul_reg_table2$Variables)
    names(mul_reg_table2)[6] <- paste0("OR(",CL*100,"%CI)")
    mul_reg_table2$`b` <- na_if(mul_reg_table2$`b`, "NA")
    mul_reg_table2$`S.E` <- na_if(mul_reg_table2$`S.E`, "NA")
    mul_reg_table2$`z` <- na_if(mul_reg_table2$`z`, "NA")
    mul_reg_table2$`exp_b_lower` <- na_if(mul_reg_table2$`exp_b_lower`, "NA") %>% as.numeric
    mul_reg_table2$`exp_b_upper` <- na_if(mul_reg_table2$`exp_b_upper`, "NA") %>% as.numeric
  }
  ####4.结果输出####
  # windowsFonts(`Times New Roman` = windowsFont("Times New Roman"))
  set_flextable_defaults(
    font.family = "Times New Roman", # 字体名称
    font.size = 10.5, # 字体大小
    border.color = "black" # 边框颜色
  )
  #####4.1单因素####
  if (reg == "uni") {
    ll <- reg_table
    first_out <- paste0("以",y,"为因变量，分别以",paste0(vars,collapse = "、"),"为自变量构建批量单因素logistic回归模型探索影响因素，结果显示：")
    i = 1
    describe_out <- vector()
    while (i <= nrow(ll)) {
      if (!is.na(ll[i,"Pr(>|z|)"])) {
        describe_out[i] <- ifelse(ll[i,"Pr(>|z|)"]<0.05,
                                  paste0(ll[i,"Variables"],"每增加1个单位，研究对象发生",y,"的可能性将增加",
                                         substr(ll[i,"OR(95%CI)"],1,regexpr("\\(",ll[i,"OR(95%CI)"])-1),
                                         "倍（OR=",substr(ll[i,"OR(95%CI)"],1,regexpr("\\(",ll[i,"OR(95%CI)"])-1),
                                         ",",CL*100,"%CI:",substr(ll[i,"OR(95%CI)"],regexpr("\\(",ll[i,"OR(95%CI)"])+1,regexpr("\\)",ll[i,"OR(95%CI)"])-1),
                                         ",p:",ll[i,"P"],"）；"),
                                  NA)
        i = i+1
      }else{
        describe_out_vec <- vector()
        for (j in 2:nlevels(data[,ll[i,"Variables"]])) {
          describe_out_vec[j] <- ifelse(ll[i+j,"Pr(>|z|)"]<0.05,
                                        paste0(ll[i,"Variables"],"为",substr(ll[i+j,"Variables"],3,nchar(ll[i+j,"Variables"])),
                                               "的研究对象发生",y,"的可能性是",
                                               ll[i,"Variables"],"为",substr(ll[i+1,"Variables"],3,nchar(ll[i+j,"Variables"])),
                                               "的研究对象的",substr(ll[i+j,"OR(95%CI)"],1,regexpr("\\(",ll[i+j,"OR(95%CI)"])-1),
                                               "倍（OR=",substr(ll[i+j,"OR(95%CI)"],1,regexpr("\\(",ll[i+j,"OR(95%CI)"])-1),
                                               ",",CL*100,"%CI:",substr(ll[i+j,"OR(95%CI)"],regexpr("\\(",ll[i+j,"OR(95%CI)"])+1,regexpr("\\)",ll[i+j,"OR(95%CI)"])-1),
                                               ",p:",ll[i+j,"P"],"）；"),
                                        NA)
          describe_out[i] <- ifelse(as.logical(prod(sapply(describe_out_vec,is.na))),NA,paste0(na.omit(describe_out_vec),collapse = "；"))
        }
        i = i+1+nlevels(data[,ll[i,"Variables"]])
      }
    }
    describe <- ifelse(length(sig_vars_description)==0,"",paste0(na.omit(describe_out),collapse = "\n"))
    final_out <- ifelse(length(vars[!vars %in% sig_vars_description])==0,"",
                        paste0(paste0(vars[!vars %in% sig_vars_description],collapse = "、"),
                               "与结局的发生可能性无关。"))

    flextable(reg_table,col_keys = names(reg_table)[-(7:9)],defaults = tibble(),theme_fun = theme_booktabs) %>%
      align(align = "left" ,part = "all") %>% #居左
      color(j = ~P,i = which(reg_table$`Pr(>|z|)` < 0.05),color = "red",part = "body") %>%
      bold(j = ~P,bold = ifelse(reg_table$`Pr(>|z|)` < 0.05, T, F), part = "body") %>%
      italic(j = ~P,part="header") %>% #P值斜体
      padding(j = ~Variables, i = which(grepl("  ", reg_table[,"Variables"])), padding.left = 20) %>% #分组留空
      add_footer_lines(values = "OR: Odds Ratio, CI: Confidence Interval") %>%
      set_table_properties(layout = "autofit") -> reg_flextable
    uni_reg_docx <- read_docx() %>%
      body_add_par("单因素回归分析结果", style = "heading 1") %>%
      body_add_par("by：student_覃泳淘", style = "centered") %>%
      body_add_par(first_out, style = "Normal") %>%
      body_add_par(describe, style = "Normal") %>%
      body_add_par(final_out, style = "Normal") %>%
      body_add_par("\t下面是单因素回归分析的表格，表格中Variables为回归模型纳入变量，b为回归系数，S.E为系数标准差，z为统计量", style = "Normal") %>%
      body_add_flextable(value = reg_flextable)
    print(uni_reg_docx, target = "单因素回归分析结果.docx")
    print(reg_flextable)
    return(reg_table)
  }
  #####4.2多因素####
  if ((reg == "mul")&is.null(direction)) {
    ll <- mul_reg_table
    first_out <- paste0("以",y,"为因变量，以",paste0(vars,collapse = "、"),"为自变量构建多因素logistic回归模型探索影响因素，结果显示：")
    i = 1
    describe_out <- vector()
    while (i <= nrow(ll)) {
      if (!is.na(ll[i,"Pr(>|z|)"])) {
        describe_out[i] <- ifelse(ll[i,"Pr(>|z|)"]<0.05,
                                  paste0(ll[i,"Variables"],"每增加1个单位，研究对象发生",y,"的可能性将增加",
                                         substr(ll[i,"OR(95%CI)"],1,regexpr("\\(",ll[i,"OR(95%CI)"])-1),
                                         "倍（OR=",substr(ll[i,"OR(95%CI)"],1,regexpr("\\(",ll[i,"OR(95%CI)"])-1),
                                         ",",CL*100,"%CI:",substr(ll[i,"OR(95%CI)"],regexpr("\\(",ll[i,"OR(95%CI)"])+1,regexpr("\\)",ll[i,"OR(95%CI)"])-1),
                                         ",p:",ll[i,"P"],"）；"),
                                  NA)
        i = i+1
      }else{
        describe_out_vec <- vector()
        for (j in 2:nlevels(data[,ll[i,"Variables"]])) {
          describe_out_vec[j] <- ifelse(ll[i+j,"Pr(>|z|)"]<0.05,
                                        paste0(ll[i,"Variables"],"为",substr(ll[i+j,"Variables"],3,nchar(ll[i+j,"Variables"])),
                                               "的研究对象发生",y,"的可能性是",
                                               ll[i,"Variables"],"为",substr(ll[i+1,"Variables"],3,nchar(ll[i+j,"Variables"])),
                                               "的研究对象的",substr(ll[i+j,"OR(95%CI)"],1,regexpr("\\(",ll[i+j,"OR(95%CI)"])-1),
                                               "倍（OR=",substr(ll[i+j,"OR(95%CI)"],1,regexpr("\\(",ll[i+j,"OR(95%CI)"])-1),
                                               ",",CL*100,"%CI:",substr(ll[i+j,"OR(95%CI)"],regexpr("\\(",ll[i+j,"OR(95%CI)"])+1,regexpr("\\)",ll[i+j,"OR(95%CI)"])-1),
                                               ",p:",ll[i+j,"P"],"）；"),
                                        NA)
          describe_out[i] <- ifelse(as.logical(prod(sapply(describe_out_vec,is.na))),NA,paste0(na.omit(describe_out_vec),collapse = "；"))
        }
        i = i+1+nlevels(data[,ll[i,"Variables"]])
      }
    }
    describe <- ifelse(length(mul_sig_vars_description)==0,"",paste0(na.omit(describe_out),collapse = "\n"))
    final_out <- ifelse(length(vars[!vars %in% mul_sig_vars_description])==0,"",
                        paste0(paste0(vars[!vars %in% mul_sig_vars_description],collapse = "、"),
                               "与结局的发生可能性无关。"))

    flextable(mul_reg_table,col_keys = names(mul_reg_table)[-(7:9)],defaults = tibble(),theme_fun = theme_booktabs) %>%
      align(align = "left" ,part = "all") %>% #居左
      color(j = ~P,i = which(mul_reg_table$`Pr(>|z|)` < 0.05),color = "red",part = "body") %>%
      bold(j = ~P,bold = ifelse(mul_reg_table$`Pr(>|z|)` < 0.05, T, F), part = "body") %>%
      italic(j = ~P,part="header") %>% #P值斜体
      padding(j = ~Variables, i = which(grepl("  ", mul_reg_table[,"Variables"])), padding.left = 20) %>% #分组留空
      add_footer_lines(values = "OR: Odds Ratio, CI: Confidence Interval") %>%
      set_table_properties(layout = "autofit") -> mul_reg_flextable
    mul_reg_docx <- read_docx() %>%
      body_add_par("多因素回归分析结果", style = "heading 1") %>%
      body_add_par("by：student_覃泳淘", style = "centered") %>%
      body_add_par(first_out, style = "Normal") %>%
      body_add_par(describe, style = "Normal") %>%
      body_add_par(final_out, style = "Normal") %>%
      body_add_par("\t下面是多因素回归分析的表格，表格中Variables为回归模型纳入变量，b为回归系数，S.E为系数标准差，z为统计量", style = "Normal") %>%
      body_add_flextable(value = mul_reg_flextable)
    print(mul_reg_docx, target = "多因素回归分析结果.docx")
    print(mul_reg_flextable)
    return(mul_reg_table)
    ll <- mul_reg_table
  }
  #####4.3多因素逐步####
  if ((reg == "mul")&(!is.null(direction))) {
    ll <- mul_reg_table2
    first_out <- paste0("以",y,"为因变量，以",paste0(vars,collapse = "、"),"为自变量构建多因素logistic逐步回归模型探索影响因素，结果显示：")
    i = 1
    describe_out <- vector()
    while (i <= nrow(ll)) {
      if (!is.na(ll[i,"Pr(>|z|)"])) {
        describe_out[i] <- ifelse(ll[i,"Pr(>|z|)"]<0.05,
                                  paste0(ll[i,"Variables"],"每增加1个单位，研究对象发生",y,"的可能性将增加",
                                         substr(ll[i,"OR(95%CI)"],1,regexpr("\\(",ll[i,"OR(95%CI)"])-1),
                                         "倍（OR=",substr(ll[i,"OR(95%CI)"],1,regexpr("\\(",ll[i,"OR(95%CI)"])-1),
                                         ",",CL*100,"%CI:",substr(ll[i,"OR(95%CI)"],regexpr("\\(",ll[i,"OR(95%CI)"])+1,regexpr("\\)",ll[i,"OR(95%CI)"])-1),
                                         ",p:",ll[i,"P"],"）；"),
                                  NA)
        i = i+1
      }else{
        describe_out_vec <- vector()
        for (j in 2:nlevels(data[,ll[i,"Variables"]])) {
          describe_out_vec[j] <- ifelse(ll[i+j,"Pr(>|z|)"]<0.05,
                                        paste0(ll[i,"Variables"],"为",substr(ll[i+j,"Variables"],3,nchar(ll[i+j,"Variables"])),
                                               "的研究对象发生",y,"的可能性是",
                                               ll[i,"Variables"],"为",substr(ll[i+1,"Variables"],3,nchar(ll[i+j,"Variables"])),
                                               "的研究对象的",substr(ll[i+j,"OR(95%CI)"],1,regexpr("\\(",ll[i+j,"OR(95%CI)"])-1),
                                               "倍（OR=",substr(ll[i+j,"OR(95%CI)"],1,regexpr("\\(",ll[i+j,"OR(95%CI)"])-1),
                                               ",",CL*100,"%CI:",substr(ll[i+j,"OR(95%CI)"],regexpr("\\(",ll[i+j,"OR(95%CI)"])+1,regexpr("\\)",ll[i+j,"OR(95%CI)"])-1),
                                               ",p:",ll[i+j,"P"],"）；"),
                                        NA)
          describe_out[i] <- ifelse(as.logical(prod(sapply(describe_out_vec,is.na))),NA,paste0(na.omit(describe_out_vec),collapse = "；"))
        }
        i = i+1+nlevels(data[,ll[i,"Variables"]])
      }
    }
    describe <- ifelse(length(mul_sig_vars_description)==0,"",paste0(na.omit(describe_out),collapse = "\n"))
    final_out <- ifelse(length(vars[!vars %in% mul_sig_vars_description])==0,"",
                        paste0(paste0(vars[!vars %in% mul_sig_vars_description],collapse = "、"),
                               "与结局的发生可能性无关。"))


    flextable(mul_reg_table2,col_keys = names(mul_reg_table2)[-(7:9)],defaults = tibble(),theme_fun = theme_booktabs) %>%
      align(align = "left" ,part = "all") %>% #居左
      color(j = ~P,i = which(mul_reg_table2$`Pr(>|z|)` < 0.05),color = "red",part = "body") %>%
      bold(j = ~P,bold = ifelse(mul_reg_table2$`Pr(>|z|)` < 0.05, T, F), part = "body") %>%
      italic(j = ~P,part="header") %>% #P值斜体
      padding(j = ~Variables, i = which(grepl("  ", mul_reg_table2[,"Variables"])), padding.left = 20) %>% #分组留空
      add_footer_lines(values = "OR: Odds Ratio, CI: Confidence Interval") %>%
      set_table_properties(layout = "autofit") -> mul_reg_flextable2
    mul_reg_docx2 <- read_docx() %>%
      body_add_par("多因素逐步回归分析结果", style = "heading 1") %>%
      body_add_par("by：student_覃泳淘", style = "centered") %>%
      body_add_par(first_out, style = "Normal") %>%
      body_add_par(describe, style = "Normal") %>%
      body_add_par(final_out, style = "Normal") %>%
      body_add_par("\t下面是多因素逐步回归分析的表格，表格中Variables为回归模型纳入变量，b为回归系数，S.E为系数标准差，z为统计量", style = "Normal") %>%
      body_add_flextable(value = mul_reg_flextable2)
    print(mul_reg_docx2, target = "多因素逐步回归分析结果.docx")
    print(mul_reg_flextable2)
    return(mul_reg_table2)
    ll <- mul_reg_table2
  }
  #####4.4单+多因素####
  if ((reg == "uni+mul")&(is.null(direction))) {
    ll <- reg_table
    first_out <- paste0("以",y,"为因变量，分别以",paste0(vars,collapse = "、"),"为自变量构建批量单因素logistic回归模型探索影响因素，结果显示：")
    i = 1
    describe_out <- vector()
    while (i <= nrow(ll)) {
      if (!is.na(ll[i,"Pr(>|z|)"])) {
        describe_out[i] <- ifelse(ll[i,"Pr(>|z|)"]<0.05,
                                  paste0(ll[i,"Variables"],"每增加1个单位，研究对象发生",y,"的可能性将增加",
                                         substr(ll[i,"OR(95%CI)"],1,regexpr("\\(",ll[i,"OR(95%CI)"])-1),
                                         "倍（OR=",substr(ll[i,"OR(95%CI)"],1,regexpr("\\(",ll[i,"OR(95%CI)"])-1),
                                         ",",CL*100,"%CI:",substr(ll[i,"OR(95%CI)"],regexpr("\\(",ll[i,"OR(95%CI)"])+1,regexpr("\\)",ll[i,"OR(95%CI)"])-1),
                                         ",p:",ll[i,"P"],"）；"),
                                  NA)
        i = i+1
      }else{
        describe_out_vec <- vector()
        for (j in 2:nlevels(data[,ll[i,"Variables"]])) {
          describe_out_vec[j] <- ifelse(ll[i+j,"Pr(>|z|)"]<0.05,
                                        paste0(ll[i,"Variables"],"为",substr(ll[i+j,"Variables"],3,nchar(ll[i+j,"Variables"])),
                                               "的研究对象发生",y,"的可能性是",
                                               ll[i,"Variables"],"为",substr(ll[i+1,"Variables"],3,nchar(ll[i+j,"Variables"])),
                                               "的研究对象的",substr(ll[i+j,"OR(95%CI)"],1,regexpr("\\(",ll[i+j,"OR(95%CI)"])-1),
                                               "倍（OR=",substr(ll[i+j,"OR(95%CI)"],1,regexpr("\\(",ll[i+j,"OR(95%CI)"])-1),
                                               ",",CL*100,"%CI:",substr(ll[i+j,"OR(95%CI)"],regexpr("\\(",ll[i+j,"OR(95%CI)"])+1,regexpr("\\)",ll[i+j,"OR(95%CI)"])-1),
                                               ",p:",ll[i+j,"P"],"）；"),
                                        NA)
          describe_out[i] <- ifelse(as.logical(prod(sapply(describe_out_vec,is.na))),NA,paste0(na.omit(describe_out_vec),collapse = "；"))
        }
        i = i+1+nlevels(data[,ll[i,"Variables"]])
      }
    }
    describe <- ifelse(length(sig_vars_description)==0,"",paste0(na.omit(describe_out),collapse = "\n"))
    final_out <- ifelse(length(vars[!vars %in% sig_vars_description])==0,"",
                        paste0(paste0(vars[!vars %in% sig_vars_description],collapse = "、"),
                               "与结局的发生可能性无关。"))
    final_describe <- paste0(first_out,na.omit(describe),na.omit(final_out),collapse = "\n")

    ll2 <- mul_reg_table
    first_out2 <- paste0("以",y,"为因变量，以",paste0(sig_vars,collapse = "、"),"为自变量构建多因素logistic回归模型探索影响因素，结果显示：")
    i = 1
    describe_out <- vector()
    while (i <= nrow(ll2)) {
      if (!is.na(ll2[i,"Pr(>|z|)"])) {
        describe_out[i] <- ifelse(ll2[i,"Pr(>|z|)"]<0.05,
                                  paste0(ll2[i,"Variables"],"每增加1个单位，研究对象发生",y,"的可能性将增加",
                                         substr(ll2[i,"OR(95%CI)"],1,regexpr("\\(",ll2[i,"OR(95%CI)"])-1),
                                         "倍（OR=",substr(ll2[i,"OR(95%CI)"],1,regexpr("\\(",ll2[i,"OR(95%CI)"])-1),
                                         ",",CL*100,"%CI:",substr(ll2[i,"OR(95%CI)"],regexpr("\\(",ll2[i,"OR(95%CI)"])+1,regexpr("\\)",ll2[i,"OR(95%CI)"])-1),
                                         ",p:",ll2[i,"P"],"）；"),
                                  NA)
        i = i+1
      }else{
        describe_out_vec <- vector()
        for (j in 2:nlevels(data[,ll2[i,"Variables"]])) {
          describe_out_vec[j] <- ifelse(ll2[i+j,"Pr(>|z|)"]<0.05,
                                        paste0(ll2[i,"Variables"],"为",substr(ll2[i+j,"Variables"],3,nchar(ll2[i+j,"Variables"])),
                                               "的研究对象发生",y,"的可能性是",
                                               ll2[i,"Variables"],"为",substr(ll2[i+1,"Variables"],3,nchar(ll2[i+j,"Variables"])),
                                               "的研究对象的",substr(ll2[i+j,"OR(95%CI)"],1,regexpr("\\(",ll2[i+j,"OR(95%CI)"])-1),
                                               "倍（OR=",substr(ll2[i+j,"OR(95%CI)"],1,regexpr("\\(",ll2[i+j,"OR(95%CI)"])-1),
                                               ",",CL*100,"%CI:",substr(ll2[i+j,"OR(95%CI)"],regexpr("\\(",ll2[i+j,"OR(95%CI)"])+1,regexpr("\\)",ll2[i+j,"OR(95%CI)"])-1),
                                               ",p:",ll2[i+j,"P"],"）；"),
                                        NA)
          describe_out[i] <- ifelse(as.logical(prod(sapply(describe_out_vec,is.na))),NA,paste0(na.omit(describe_out_vec),collapse = "；"))
        }
        i = i+1+nlevels(data[,ll2[i,"Variables"]])
      }
    }
    describe2 <- ifelse(length(mul_sig_vars_description)==0,"",paste0(na.omit(describe_out),collapse = "\n"))
    final_out2 <- ifelse(length(sig_vars[!sig_vars %in% mul_sig_vars_description])==0,"",
                        paste0(paste0(sig_vars[!sig_vars %in% mul_sig_vars_description],collapse = "、"),
                               "与结局的发生可能性无关。"))


    if(!require("gridExtra")) install.packages("gridExtra")
    library(gridExtra)



    flextable(reg_table,col_keys = names(reg_table)[-(7:9)],defaults = tibble(),theme_fun = theme_booktabs) %>%
      align(align = "left" ,part = "all") %>% #居左
      color(j = ~P,i = which(reg_table$`Pr(>|z|)` < 0.05),color = "red",part = "body") %>%
      bold(j = ~P,bold = ifelse(reg_table$`Pr(>|z|)` < 0.05, T, F), part = "body") %>%
      italic(j = ~P,part="header") %>% #P值斜体
      padding(j = ~Variables, i = which(grepl("  ", reg_table[,"Variables"])), padding.left = 20) %>% #分组留空
      add_footer_lines(values = "OR: Odds Ratio, CI: Confidence Interval") %>%
      set_table_properties(layout = "autofit") -> reg_flextable

    flextable(mul_reg_table,col_keys = names(mul_reg_table)[-(7:9)],defaults = tibble(),theme_fun = theme_booktabs) %>%
      align(align = "left" ,part = "all") %>% #居左
      color(j = ~P,i = which(mul_reg_table$`Pr(>|z|)` < 0.05),color = "red",part = "body") %>%
      bold(j = ~P,bold = ifelse(mul_reg_table$`Pr(>|z|)` < 0.05, T, F), part = "body") %>%
      italic(j = ~P,part="header") %>% #P值斜体
      padding(j = ~Variables, i = which(grepl("  ", mul_reg_table[,"Variables"])), padding.left = 20) %>% #分组留空
      add_footer_lines(values = "OR: Odds Ratio, CI: Confidence Interval") %>%
      set_table_properties(layout = "autofit") -> mul_reg_flextable
    mulreg_part <- block_section(prop_section(page_size = page_size(orient = "portrait"), type = "continuous"))

    `uni+mul_reg_table_docx` <- read_docx() %>%
      body_add_par("单因素回归分析结果", style = "heading 1") %>%
      body_add_par("by：student_覃泳淘", style = "centered") %>%
      body_add_par(first_out, style = "Normal") %>%
      body_add_par(describe, style = "Normal") %>%
      body_add_par(final_out, style = "Normal") %>%
      body_add_par("\t下面是单因素回归分析的表格，表格中Variables为回归模型纳入变量，b为回归系数，S.E为系数标准差，z为统计量", style = "Normal") %>%
      body_add_flextable(value = reg_flextable) %>%
      body_end_block_section(value = mulreg_part) %>%
      body_add_par("多因素回归分析结果", style = "heading 1") %>%
      body_add_par("覃泳淘", style = "centered") %>%
      body_add_par(first_out2, style = "Normal") %>%
      body_add_par(describe2, style = "Normal") %>%
      body_add_par(final_out2, style = "Normal") %>%
      body_add_par("\t下面是多因素回归分析的表格，表格中Variables为回归模型纳入变量，b为回归系数，S.E为系数标准差，z为统计量", style = "Normal") %>%
      body_add_flextable(value = mul_reg_flextable)
    print(`uni+mul_reg_table_docx`, target = "单+多因素回归分析结果.docx")
    print(reg_flextable)
    print(mul_reg_flextable)
    `sin+mul_reg_list` <- list(reg_table,mul_reg_table)
    return(`sin+mul_reg_list`)
  }



  #####4.5单+多因素逐步####
  if ((reg == "uni+mul")&(!is.null(direction))) {
    ll <- reg_table
    first_out <- paste0("以",y,"为因变量，分别以",paste0(vars,collapse = "、"),"为自变量构建批量单因素logistic回归模型探索影响因素，结果显示：")
    i = 1
    describe_out <- vector()
    while (i <= nrow(ll)) {
      if (!is.na(ll[i,"Pr(>|z|)"])) {
        describe_out[i] <- ifelse(ll[i,"Pr(>|z|)"]<0.05,
                                  paste0(ll[i,"Variables"],"每增加1个单位，研究对象发生",y,"的可能性将增加",
                                         substr(ll[i,"OR(95%CI)"],1,regexpr("\\(",ll[i,"OR(95%CI)"])-1),
                                         "倍（OR=",substr(ll[i,"OR(95%CI)"],1,regexpr("\\(",ll[i,"OR(95%CI)"])-1),
                                         ",",CL*100,"%CI:",substr(ll[i,"OR(95%CI)"],regexpr("\\(",ll[i,"OR(95%CI)"])+1,regexpr("\\)",ll[i,"OR(95%CI)"])-1),
                                         ",p:",ll[i,"P"],"）；"),
                                  NA)
        i = i+1
      }else{
        describe_out_vec <- vector()
        for (j in 2:nlevels(data[,ll[i,"Variables"]])) {
          describe_out_vec[j] <- ifelse(ll[i+j,"Pr(>|z|)"]<0.05,
                                        paste0(ll[i,"Variables"],"为",substr(ll[i+j,"Variables"],3,nchar(ll[i+j,"Variables"])),
                                               "的研究对象发生",y,"的可能性是",
                                               ll[i,"Variables"],"为",substr(ll[i+1,"Variables"],3,nchar(ll[i+j,"Variables"])),
                                               "的研究对象的",substr(ll[i+j,"OR(95%CI)"],1,regexpr("\\(",ll[i+j,"OR(95%CI)"])-1),
                                               "倍（OR=",substr(ll[i+j,"OR(95%CI)"],1,regexpr("\\(",ll[i+j,"OR(95%CI)"])-1),
                                               ",",CL*100,"%CI:",substr(ll[i+j,"OR(95%CI)"],regexpr("\\(",ll[i+j,"OR(95%CI)"])+1,regexpr("\\)",ll[i+j,"OR(95%CI)"])-1),
                                               ",p:",ll[i+j,"P"],"）；"),
                                        NA)
          describe_out[i] <- ifelse(as.logical(prod(sapply(describe_out_vec,is.na))),NA,paste0(na.omit(describe_out_vec),collapse = "；"))
        }
        i = i+1+nlevels(data[,ll[i,"Variables"]])
      }
    }
    describe <- ifelse(length(sig_vars_description)==0,"",paste0(na.omit(describe_out),collapse = "\n"))
    final_out <- ifelse(length(vars[!vars %in% sig_vars_description])==0,"",
                        paste0(paste0(vars[!vars %in% sig_vars_description],collapse = "、"),
                               "与结局的发生可能性无关。"))
    final_describe <- paste0(first_out,na.omit(describe),na.omit(final_out),collapse = "\n")

    ll2 <- mul_reg_table2
    first_out2 <- paste0("以",y,"为因变量，以",paste0(sig_vars,collapse = "、"),"为自变量构建多因素logistic逐步回归模型探索影响因素，结果显示：")
    i = 1
    describe_out <- vector()
    while (i <= nrow(ll2)) {
      if (!is.na(ll2[i,"Pr(>|z|)"])) {
        describe_out[i] <- ifelse(ll2[i,"Pr(>|z|)"]<0.05,
                                  paste0(ll2[i,"Variables"],"每增加1个单位，研究对象发生",y,"的可能性将增加",
                                         substr(ll2[i,"OR(95%CI)"],1,regexpr("\\(",ll2[i,"OR(95%CI)"])-1),
                                         "倍（OR=",substr(ll2[i,"OR(95%CI)"],1,regexpr("\\(",ll2[i,"OR(95%CI)"])-1),
                                         ",",CL*100,"%CI:",substr(ll2[i,"OR(95%CI)"],regexpr("\\(",ll2[i,"OR(95%CI)"])+1,regexpr("\\)",ll2[i,"OR(95%CI)"])-1),
                                         ",p:",ll2[i,"P"],"）；"),
                                  NA)
        i = i+1
      }else{
        describe_out_vec <- vector()
        for (j in 2:nlevels(data[,ll2[i,"Variables"]])) {
          describe_out_vec[j] <- ifelse(ll2[i+j,"Pr(>|z|)"]<0.05,
                                        paste0(ll2[i,"Variables"],"为",substr(ll2[i+j,"Variables"],3,nchar(ll2[i+j,"Variables"])),
                                               "的研究对象发生",y,"的可能性是",
                                               ll2[i,"Variables"],"为",substr(ll2[i+1,"Variables"],3,nchar(ll2[i+j,"Variables"])),
                                               "的研究对象的",substr(ll2[i+j,"OR(95%CI)"],1,regexpr("\\(",ll2[i+j,"OR(95%CI)"])-1),
                                               "倍（OR=",substr(ll2[i+j,"OR(95%CI)"],1,regexpr("\\(",ll2[i+j,"OR(95%CI)"])-1),
                                               ",",CL*100,"%CI:",substr(ll2[i+j,"OR(95%CI)"],regexpr("\\(",ll2[i+j,"OR(95%CI)"])+1,regexpr("\\)",ll2[i+j,"OR(95%CI)"])-1),
                                               ",p:",ll2[i+j,"P"],"）；"),
                                        NA)
          describe_out[i] <- ifelse(as.logical(prod(sapply(describe_out_vec,is.na))),NA,paste0(na.omit(describe_out_vec),collapse = "；"))
        }
        i = i+1+nlevels(data[,ll2[i,"Variables"]])
      }
    }
    describe2 <- ifelse(length(mul_sig_vars_description)==0,"",paste0(na.omit(describe_out),collapse = "\n"))
    final_out2 <- ifelse(length(sig_vars[!sig_vars %in% mul_sig_vars_description])==0,"",
                        paste0(paste0(sig_vars[!sig_vars %in% mul_sig_vars_description],collapse = "、"),
                               "与结局的发生可能性无关。"))





    if(!require("gridExtra")) install.packages("gridExtra")
    library(gridExtra)
    flextable(reg_table,col_keys = names(reg_table)[-(7:9)],defaults = tibble(),theme_fun = theme_booktabs) %>%
      align(align = "left" ,part = "all") %>% #居左
      color(j = ~P,i = which(reg_table$`Pr(>|z|)` < 0.05),color = "red",part = "body") %>%
      bold(j = ~P,bold = ifelse(reg_table$`Pr(>|z|)` < 0.05, T, F), part = "body") %>%
      italic(j = ~P,part="header") %>% #P值斜体
      padding(j = ~Variables, i = which(grepl("  ", reg_table[,"Variables"])), padding.left = 20) %>% #分组留空
      add_footer_lines(values = "OR: Odds Ratio, CI: Confidence Interval") %>%
      set_table_properties(layout = "autofit") -> reg_flextable

    flextable(mul_reg_table2,col_keys = names(mul_reg_table2)[-(7:9)],defaults = tibble(),theme_fun = theme_booktabs) %>%
      align(align = "left" ,part = "all") %>% #居左
      color(j = ~P,i = which(mul_reg_table2$`Pr(>|z|)` < 0.05),color = "red",part = "body") %>%
      bold(j = ~P,bold = ifelse(mul_reg_table2$`Pr(>|z|)` < 0.05, T, F), part = "body") %>%
      italic(j = ~P,part="header") %>% #P值斜体
      padding(j = ~Variables, i = which(grepl("  ", mul_reg_table2[,"Variables"])), padding.left = 20) %>% #分组留空
      add_footer_lines(values = "OR: Odds Ratio, CI: Confidence Interval") %>%
      set_table_properties(layout = "autofit") -> mul_reg_flextable
    mulreg_part <- block_section(prop_section(page_size = page_size(orient = "portrait"), type = "continuous"))

    `uni+mul_reg_table2_docx` <- read_docx() %>%
      body_add_par("单因素回归分析结果", style = "heading 1") %>%
      body_add_par("by：student_覃泳淘", style = "centered") %>%
      body_add_par(first_out, style = "Normal") %>%
      body_add_par(describe, style = "Normal") %>%
      body_add_par(final_out, style = "Normal") %>%
      body_add_par("\t下面是单因素回归分析的表格，表格中Variables为回归模型纳入变量，b为回归系数，S.E为系数标准差，z为统计量", style = "Normal") %>%
      body_add_flextable(value = reg_flextable) %>%
      body_end_block_section(value = mulreg_part) %>%
      body_add_par("多因素逐步回归分析结果", style = "heading 1") %>%
      body_add_par("覃泳淘", style = "centered") %>%
      body_add_par(first_out2, style = "Normal") %>%
      body_add_par(describe2, style = "Normal") %>%
      body_add_par(final_out2, style = "Normal") %>%
      body_add_par("\t下面是多因素逐步回归分析的表格，表格中Variables为回归模型纳入变量，b为回归系数，S.E为系数标准差，z为统计量", style = "Normal") %>%
      body_add_flextable(value = mul_reg_flextable)
    print(`uni+mul_reg_table2_docx`, target = "单+多因素逐步回归分析结果.docx")
    print(reg_flextable)
    print(mul_reg_flextable)
    `sin+mul_reg_list` <- list(reg_table,mul_reg_table2)
    return(`sin+mul_reg_list`)
  }



}

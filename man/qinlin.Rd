% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/qinlin.R
\name{qinlin}
\alias{qinlin}
\title{单+多因素线性回归回归}
\usage{
qinlin(
  data = NULL,
  vars = NULL,
  y = NULL,
  reg = NULL,
  decimal = NULL,
  P_decimal = NULL,
  sig_P = NULL,
  CL = NULL,
  direction = NULL
)
}
\arguments{
\item{data}{*输入数据框}

\item{vars}{*自变量}

\item{y}{*结局变量}

\item{reg}{回归方式："uni","mul","uni+mul"（默认）}

\item{decimal}{结果保留小数点位，默认为3}

\item{P_decimal}{P值保留小数点位，默认为3}

\item{sig_P}{多因素回归纳入变量P值设定，默认为0.05}

\item{CL}{置信度，默认为0.95}

\item{direction}{逐步回归方向："both", "backward", "forward",参数缺失则不进行逐步回归}
}
\description{
该函数可用于进行批量单因素线性回归分析以及多因素线性回归分析。
}
\author{
qinstudent
}

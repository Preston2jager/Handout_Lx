\LoadClass{article}
%=========================================================
\usepackage{graphicx}
\usepackage{fancyhdr}
\usepackage{tabularx}
\usepackage{longtable}
\usepackage{geometry}
\usepackage{fontspec} 
\usepackage{enumerate}
\usepackage{amssymb}
%=========================================================
%For font
\setmainfont{Times New Roman}
%=========================================================
%For page margin
\geometry{a4paper,left=2cm,right=2cm,top=2cm,bottom=3cm}
%=========================================================
\pagestyle{fancy}
\renewcommand{\headrulewidth}{0pt} % 去页眉线
\fancyhead{} % 页眉清空
\fancyfoot{} % 页脚清空
\fancyhead[R]{\course}
\fancyhead[L]{
	\begin{minipage}[c]{0.06\textwidth}
		\includegraphics[height=8mm]{./Resources/logo/Logo.png}
	\end{minipage}
}
\fancyfoot[L]{\week~\episode}
\fancyfoot[R]{Page \thepage}
%=========================================================
\newcounter{steps}
\newcommand{\steplist}{\refstepcounter{steps} \thesteps}

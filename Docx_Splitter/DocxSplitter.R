

# 加载必要的包
library(shiny)
library(shinyFiles)
library(officer)

# 定义函数：提取 .docx 文件内容并按行保存
extract_and_save_docx <- function(docx_file, output_folder) {
  tryCatch({
    # 读取 .docx 文件
    doc <- read_docx(docx_file)
    doc_summary <- docx_summary(doc)
    text <- doc_summary$text
    
    # 创建以文件名命名的子文件夹
    file_name <- tools::file_path_sans_ext(basename(docx_file))  # 去除扩展名
    subfolder <- file.path(output_folder, file_name)  # 子文件夹路径
    if (!dir.exists(subfolder)) {
      dir.create(subfolder, recursive = TRUE)  # 如果子文件夹不存在，则创建
    }
    
    # 按行拆分内容并保存为单独的 .txt 文件
    for (i in 1:length(text)) {
      line <- text[i]
      output_file <- file.path(subfolder, paste0("line_", i, ".txt"))  # 生成输出文件路径
      writeLines(line, output_file)  # 将每一行写入单独的 .txt 文件
    }
    
    cat("文件已拆分并保存到:", subfolder, "\n")
  }, error = function(e) {
    cat("处理文件时出错:", docx_file, "\n")
    cat("错误信息:", e$message, "\n")
  })
}

# 定义函数：批量处理文件夹中的所有 .docx 文件
batch_process_docx <- function(input_folder, output_folder) {
  # 获取所有 .docx 文件
  docx_files <- list.files(input_folder, pattern = "\\.docx$", full.names = TRUE)
  if (length(docx_files) == 0) {
    cat("未找到任何 .docx 文件，请检查输入文件夹路径。\n")
    return()
  }
  
  # 创建输出主文件夹
  if (!dir.exists(output_folder)) {
    dir.create(output_folder, recursive = TRUE)
  }
  
  # 遍历每个 .docx 文件并提取内容
  total_files <- length(docx_files)
  cat("共发现", total_files, "个 .docx 文件\n")
  for (i in seq_along(docx_files)) {
    cat("正在处理第", i, "个文件（共", total_files, "个）:", docx_files[i], "\n")
    extract_and_save_docx(docx_files[i], output_folder)
  }
  
  cat("批量处理完成，所有文件已保存到:", output_folder, "\n")
}

# 定义 Shiny 用户界面
ui <- fluidPage(
  titlePanel("DOCX 文件拆分工具"),
  sidebarLayout(
    sidebarPanel(
      radioButtons("mode", "请选择模式：",
                   choices = c("批量上传文件夹", "单文件上传"),
                   selected = "批量上传文件夹"),
      conditionalPanel(
        condition = "input.mode == '批量上传文件夹'",
        shinyDirButton("folder", "选择文件夹", "请选择包含 .docx 文件的文件夹")
      ),
      conditionalPanel(
        condition = "input.mode == '单文件上传'",
        fileInput("file", "请选择 .docx 文件", accept = ".docx")
      ),
      textInput("output_folder", "请输入输出文件夹路径", value = "C:/output"),
      actionButton("run", "开始处理")
    ),
    mainPanel(
      verbatimTextOutput("status")
    )
  )
)

# 定义 Shiny 服务器逻辑
server <- function(input, output, session) {
  # 设置文件夹选择
  volumes <- c(Home = fs::path_home(), "R Installation" = R.home(), getVolumes()())
  shinyDirChoose(input, "folder", roots = volumes, session = session)
  
  observeEvent(input$run, {
    output_folder <- input$output_folder
    if (input$mode == "批量上传文件夹") {
      folder_path <- parseDirPath(volumes, input$folder)
      if (length(folder_path) == 0) {
        output$status <- renderText("请选择文件夹！")
        return()
      }
      batch_process_docx(folder_path, output_folder)
      output$status <- renderText(paste("批量处理完成，文件已保存到:", output_folder))
    } else {
      docx_file <- input$file$datapath
      if (is.null(docx_file)) {
        output$status <- renderText("请选择文件！")
        return()
      }
      extract_and_save_docx(docx_file, output_folder)
      output$status <- renderText(paste("文件已拆分并保存到:", output_folder))
    }
  })
}

# 运行 Shiny 应用
shinyApp(ui = ui, server = server)
docx_replace_embedded_xlsx <- function(doc, rId, chartData, xAxisCols) {
  relationshipDT = doc$doc_obj$relationship()$get_data()
  setDT(relationshipDT)
  charts_dir = file.path(doc$package_dir, "word/charts")
  
  ################################################
  # REPLACE EMBEDDED EXCEL ASSOCIATED TO THE CHART rId
  
  chartData = cbind(chartData[,xAxisCols,with=FALSE],chartData[,-xAxisCols,with=FALSE])
  
  rel_filename = file.path(charts_dir, "_rels", paste0(strsplit(relationshipDT[id==eval(rId),target],"/")[[1]][2], ".rels"))
  rel_content = read_xml(rel_filename)
  attributes = xml_attrs(xml_children(rel_content)[[1]])
  xlsx_dir = file.path(doc$package_dir, "word/embeddings")
  dir.create(xlsx_dir, showWarnings = FALSE)
  if(attributes["Type"]=="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package") {
    xlsx_file = strsplit(attributes["Target"],"/")[[1]][3]
	# xlsl::write.xlsx2(chartData, file.path(xlsx_dir, xlsx_file), sheetName="sheet1", row.names=FALSE)
	writexl::write_xlsx(x = list(sheet1 = chartData), path = file.path(xlsx_dir, xlsx_file))
	
	# wb <- openxlsx::createWorkbook()
	# openxlsx::addWorksheet(wb, "sheet1")
	# openxlsx::writeData(wb, sheet = "sheet1", x = chartData)
	# openxlsx::saveWorkbook(wb, file.path(xlsx_dir, xlsx_file), overwrite = TRUE)
	
  } else if(attributes["Type"]=="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject") {
    chart_xml = xml_children(rel_content)[[1]]
    xlsx_path = tempfile(tmpdir = xlsx_dir, pattern = "data", fileext = ".xlsx")
    xml_attr(chart_xml, "Type") = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package"
	# xlsl::write.xlsx2(chartData, xlsx_path, sheetName = "sheet1", row.names = FALSE)
	writexl::write_xlsx(x = list(sheet1 = chartData), path = xlsx_path)
	
	# wb <- openxlsx::createWorkbook()
	# openxlsx::addWorksheet(wb, "sheet1")
	# openxlsx::writeData(wb, sheet = "sheet1", x = chartData)
	# openxlsx::saveWorkbook(wb, xlsx_path, overwrite = TRUE)
	
    xlsx_file = basename(xlsx_path)
    xml_attr(chart_xml, "Target") = paste0("../embeddings/",xlsx_file)
    xml_attr(chart_xml, "TargetMode") = NULL
    attributes = xml_attrs(xml_children(rel_content)[[1]])
    write_xml(rel_content, rel_filename)
  }
  
  # override = setNames("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", paste0("/word/embeddings/",xlsx_file))
  # doc$content_type$add_override(value = override)
     
  doc
}

docx_update_series_chart <- function(doc, rId, chartData, xAxisCols) {
  relationshipDT = doc$doc_obj$relationship()$get_data()
  setDT(relationshipDT)
  charts_dir = file.path(doc$package_dir, "word/charts")
  
  ################################################
  # UPDATE THE SERIES IN CHART rId XML
  
  chart_xml_path = file.path(charts_dir, strsplit(relationshipDT[id==eval(rId),target],"/")[[1]][2])
  chart_xml_content = read_xml(chart_xml_path)

  series = xml_find_all(chart_xml_content, "//c:ser")
  
  OldNumSeries = length(xml_length(series))
  
  isBubbleChart = ifelse(length(xml_find_all(chart_xml_content, "//c:bubbleChart"))>0,TRUE,FALSE)
  isScatterChart = ifelse(length(xml_find_all(chart_xml_content, "//c:scatterChart"))>0,TRUE,FALSE)
  if(isBubbleChart) {
    NewNumSeries = as.integer((ncol(chartData) - length(xAxisCols))/2L)
  } else {
    NewNumSeries = ncol(chartData) - length(xAxisCols)
  }
  
  diffNumSeries = NewNumSeries - OldNumSeries
  if(diffNumSeries<0) {
    for(i in (OldNumSeries+diffNumSeries+1):OldNumSeries) {
      xml_remove(series[i])
    }
  } else if(diffNumSeries>0){
    for(i in 1:abs(diffNumSeries)) {
      xml_add_sibling(series[OldNumSeries+(i-1)], series[OldNumSeries], .where="after", .copy=TRUE)
    }
    series = xml_find_all(chart_xml_content, "//c:ser")
    for(i in (OldNumSeries+1):NewNumSeries) {
      seriesChildren = xml_children(series[i])
      lengthSeriesChildren = length(xml_length(seriesChildren))
      for(j in 1:lengthSeriesChildren) {
        if(xml_name(seriesChildren[j])%in%c("idx")) {
          xml_attr(seriesChildren[j], "val") = i-1
        } else if(xml_name(seriesChildren[j])%in%c("order")) {
          xml_attr(seriesChildren[j], "val") = i-1
        }
      }
    }
  }
  
  serieNameRange = xml_find_all(chart_xml_content, "//c:tx/c:strRef/c:f")
  for(i in 1:NewNumSeries) {
    xml_text(serieNameRange[i]) = paste0("sheet1!$",toupper(letters[i+length(xAxisCols)]),"$1")
  }
  
  serieNameCache = xml_find_all(chart_xml_content, "//c:strCache")
  lengthSerieNameCache = length(xml_length(serieNameCache))
  for(i in 1:lengthSerieNameCache) {
    xml_remove(serieNameCache[i])
  }
  
  if(isScatterChart) {
    xValRange = xml_find_all(chart_xml_content, "//c:xVal/c:numRef/c:f")
    for(i in 1:NewNumSeries) {
      xml_text(xValRange[i]) = paste0("sheet1!$A$2:$A$",nrow(chartData)+1)
    }
    
    yValRange = xml_find_all(chart_xml_content, "//c:yVal/c:numRef/c:f")
    for(i in 1:NewNumSeries) {
      xml_text(yValRange[i]) = paste0("sheet1!$",toupper(letters[i+1]),"$2:$",toupper(letters[i+1]),"$",nrow(chartData)+1)
    }
  } else if(isBubbleChart) {
    xValRange = xml_find_all(chart_xml_content, "//c:xVal/c:numRef/c:f")
    for(i in 1:NewNumSeries) {
      xml_text(xValRange[i]) = paste0("sheet1!$A$2:$A$",nrow(chartData)+1)
    }
    
    yValRange = xml_find_all(chart_xml_content, "//c:yVal/c:numRef/c:f")
    for(i in 1:NewNumSeries) {
      xml_text(yValRange[i]) = paste0("sheet1!$",toupper(letters[(2*i)]),"$2:$",toupper(letters[(2*i)]),"$",nrow(chartData)+1)
    }
    
    bubbleSizeRange = xml_find_all(chart_xml_content, "//c:bubbleSize/c:numRef/c:f")
    for(i in 1:NewNumSeries) {
      xml_text(bubbleSizeRange[i]) = paste0("sheet1!$",toupper(letters[(2*i)+1]),"$2:$",toupper(letters[(2*i)+1]),"$",nrow(chartData)+1)
    }
  } else {
    serieCatRange = xml_find_all(chart_xml_content, "//c:cat/c:numRef/c:f")
    if(length(serieCatRange)==0) {
      serieCatRange = xml_find_all(chart_xml_content, "//c:cat/c:strRef/c:f")
	  if(length(serieCatRange)==0) {
        serieCatRange = xml_find_all(chart_xml_content, "//c:cat/c:multiLvlStrRef/c:f")
      }
	}
    
    for(i in 1:NewNumSeries) {
      xml_text(serieCatRange[i]) = paste0("sheet1!$A$2:$",toupper(letters[length(xAxisCols)]),"$",nrow(chartData)+1)
    }
    
    serieValNumRange = xml_find_all(chart_xml_content, "//c:val/c:numRef/c:f")
    for(i in 1:NewNumSeries) {
      xml_text(serieValNumRange[i]) = paste0("sheet1!$",toupper(letters[i+length(xAxisCols)]),"$2:$",toupper(letters[i+length(xAxisCols)]),"$",nrow(chartData)+1)
    }
  }
  
  serieXYAxisCache = xml_find_all(chart_xml_content, "//c:numCache")
  lengthSerieXYAxisCache = length(xml_length(serieXYAxisCache))
  for(i in 1:lengthSerieXYAxisCache) {
    xml_remove(serieXYAxisCache[i])
  }
  
  multiLvlStrCache = xml_find_all(chart_xml_content, "//c:multiLvlStrCache")
  lengthMultiLvlStrCache = length(xml_length(multiLvlStrCache))
  for(i in 1:lengthMultiLvlStrCache) {
    xml_remove(multiLvlStrCache[i])
  }
  
  extLst = xml_find_all(chart_xml_content, "//c:extLst")
  lengthExtLst = length(xml_length(extLst))
  for(i in 1:lengthExtLst) {
    xml_remove(extLst[i])
  }
  
  write_xml(chart_xml_content, chart_xml_path)
  
  doc
}

#' @export
#' @title update chart data from a Word document
#' @description update a \code{ms_chart} from an rdocx object, the data chart will be updated.
#' @param x an rdocx object.
#' @param rId the relation Id of the \code{ms_chart} object already in the rdocx object.
#' @param chartData data.table object containing the new chart data.
#' @param xAxisCols an integer array containing the column indexes which represent the x axis values.
#' @examples
#' library(data.table)
#' library(officer)
#' docx = read_docx(path = "PACKAGE/resource/test-docx-office-chart_test1.docx")
#' res = body_get_chart_rid_title(docx)
#' rId = res[startsWith(chartPreviousParagraph, "Chart 1 -")]$rId
#'
#' \donttest{
#' categoricalcolumn1 = c(1990, 1991, 1992)
#' valuecolumn1 = c(6, 78, 12)
#' valuecolumn2 = c(5, 1, 50)
#' df = data.table(categoricalcolumn1, valuecolumn1, valuecolumn2)
#' docx = docx_update_data_chart(x=docx, rId=rId, chartData=df, xAxisCols=c(1L))
#' docm = docx_add_vba_macro_refresh_charts_before_opening(docx)
#' print(docm, target = "result-test-docx-office-chart_test1.docm")
#' }
docx_update_data_chart <- function(x, rId, chartData, xAxisCols=c(1L)) {
  if(!(is.character(rId) && length(rId)>0))
    stop("rId must be a string", call. = FALSE)
  
  if(!is.data.table(chartData))
    stop("chartData must be a data.table object", call. = FALSE)
  
  if(any(!is.integer(xAxisCols)))
    stop("xAxisCols must be an integer array", call. = FALSE)
  
  if(length(xAxisCols)>=ncol(chartData))
    stop("xAxisCols must be an integer array", call. = FALSE)
  
  x = docx_replace_embedded_xlsx(x, rId, chartData, xAxisCols)
  
  x = docx_update_series_chart(x, rId, chartData, xAxisCols)
  
  x
}
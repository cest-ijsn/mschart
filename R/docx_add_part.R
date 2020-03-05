docx_add_vba_macro_refresh_charts_before_opening <- function(doc) {
  ################################################
  # AS ALL CACHE WERE REMOVED, INSERT VBA MACRO TO UPDATE THE CHART LOOK DURING THE .DOCM OPENING
  word_dir = file.path(doc$package_dir, "word")
  if (!file.exists(paste0(word_dir, "/vbaProject.bin"))) {  
	  fileList=c(system.file(package = "mschart", "template", "vbaData.xml"),
				 system.file(package = "mschart", "template", "vbaProject.bin"))
	  file.copy(fileList, word_dir)
	  
	  vbarel_filename = file.path(word_dir, "_rels", "vbaProject.bin.rels")
	  vbarel_str = paste0("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>",
						  "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.microsoft.com/office/2006/relationships/wordVbaData\" Target=\"vbaData.xml\"/></Relationships>")
	  cat(vbarel_str, file = vbarel_filename)
	  
	  next_id = doc$doc_obj$relationship()$get_next_id()
	  doc$doc_obj$relationship()$add(
		paste0("rId", next_id),
		type = "http://schemas.microsoft.com/office/2006/relationships/vbaProject",
		target = "vbaProject.bin"
	  )
	  
	  override = setNames("application/vnd.ms-word.vbaData+xml", "/word/vbaData.xml")
	  doc$content_type$add_override(value = override)
	  
	  # override = setNames("application/vnd.ms-office.vbaProject", "/word/vbaProject.bin")
	  # doc$content_type$add_override(value = override)
	  
	  doc$content_type$add_ext(type="application/vnd.ms-office.vbaProject", extension="bin")	  
	  doc$content_type$add_ext(type="application/octet-stream", extension="xlsx")
	  
	  doc$content_type$remove_slide("/word/document.xml")
	  override = setNames("application/vnd.ms-word.document.macroEnabled.main+xml", "/word/document.xml")
	  doc$content_type$add_override(value = override)
	  
	  doc
  } else {
	stop("This method should be called only once and preferably before the final version be saved.")
  }
}
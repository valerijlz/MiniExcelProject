// Надежное сохранение данных в рабочий файл и финальный выгруз в Uri пользователя
    private fun commitChangesAndExportToOriginal(jsonData: String) {
        val targetUri = currentFileUri
        val currentWorkingFile = workingFile ?: return

        lifecycleScope.launch {
            try {
                withContext(Dispatchers.IO) {
                    val root = JSONObject(jsonData)
                    val matrix = root.optJSONArray("matrix") ?: JSONArray()
                    val isFinalSave = root.optBoolean("isFinalSave", false)

                    // 1. Открываем рабочую копию через POI
                    val fileInputStream = FileInputStream(currentWorkingFile)
                    val workbook = WorkbookFactory.create(fileInputStream)
                    fileInputStream.close()

                    val sheet = workbook.getSheetAt(0)

                    for (r in 0 until matrix.length()) {
                        val rowArray = matrix.optJSONArray(r) ?: continue
                        var poiRow = sheet.getRow(r)
                        if (poiRow == null) {
                            poiRow = sheet.createRow(r)
                        }

                        for (c in 0 until rowArray.length()) {
                            val cellObj = rowArray.optJSONObject(c) ?: continue
                            val cellValue = cellObj.opt("v") ?: ""

                            var poiCell = poiRow.getCell(c)
                            if (poiCell == null) {
                                poiCell = poiRow.createCell(c)
                            }

                            when (cellValue) {
                                is Number -> poiCell.setCellValue(cellValue.toDouble())
                                is Boolean -> poiCell.setCellValue(cellValue)
                                else -> poiCell.setCellValue(cellValue.toString())
                            }
                        }
                    }

                    // 2. Перезаписываем временный рабочий файл
                    FileOutputStream(currentWorkingFile).use { out ->
                        workbook.write(out)
                    }
                    workbook.close()

                    // 3. Если это финальное сохранение (нажата кнопка) — выгружаем в исходный Uri
                    if (isFinalSave && targetUri != null) {
                        contentResolver.openOutputStream(targetUri, "w")?.use { outStream ->
                            currentWorkingFile.inputStream().use { inStream ->
                                inStream.copyTo(outStream)
                            }
                        }
                    }
                }
                
                // Выводим уведомление только при реальном нажатии на кнопку «Сохранить»
                val rootCheck = JSONObject(jsonData)
                if (rootCheck.optBoolean("isFinalSave", false)) {
                    withContext(Dispatchers.Main) {
                        Toast.makeText(this@MainActivity, "Файл успешно сохранен", Toast.LENGTH_SHORT).show()
                    }
                }
            } catch (e: Exception) {
                Log.e("MiniExcelDebug", "Ошибка записи: ${e.message}", e)
                val rootCheck = JSONObject(jsonData)
                if (rootCheck.optBoolean("isFinalSave", false)) {
                    withContext(Dispatchers.Main) {
                        Toast.makeText(this@MainActivity, "Ошибка сохранения: ${e.localizedMessage}", Toast.LENGTH_LONG).show()
                    }
                }
            }
        }
    }

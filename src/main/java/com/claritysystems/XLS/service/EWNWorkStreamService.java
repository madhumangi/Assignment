package com.claritysystems.XLS.service;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONArray;
import org.json.JSONObject;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.stream.IntStream;
import java.util.stream.Stream;

@Service
public class EWNWorkStreamService {

    public byte[] processXlsFile(MultipartFile file) throws IOException {
        XSSFWorkbook workbook = null;
        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();

        try {
            workbook = new XSSFWorkbook(file.getInputStream());
            Sheet sheet = workbook.getSheetAt(0);
            processSheet(sheet);
            workbook.write(outputStream);
            return outputStream.toByteArray();
        } finally {
            if (workbook != null) {
                workbook.close();
            }
        }
    }

    private void processSheet(Sheet sheet) {
        Row headerRow = sheet.getRow(0);
        int rulesColumnIndex = headerRow.getLastCellNum();
        CellStyle headerStyle = headerRow.getCell(1).getCellStyle();
        Cell rulesHeaderCell = headerRow.createCell(rulesColumnIndex);
        rulesHeaderCell.setCellValue("Rules");
        rulesHeaderCell.setCellStyle(headerStyle);

        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            if (row == null || row.getCell(1) == null || row.getCell(1).getStringCellValue().trim().isEmpty()) {
                continue; // Skip empty rows
            }

            Cell requiredTasksCell = row.getCell(1);
            String requiredTasks = requiredTasksCell.getStringCellValue();
            String rulesJson = convertToRulesJson(requiredTasks);
            row.createCell(rulesColumnIndex).setCellValue(rulesJson);
        }
    }


    private String convertToRulesJson(String requiredTasks) {
        if (requiredTasks == null || requiredTasks.trim().isEmpty()) {
            return null; // Skip empty rows
        }

        boolean hasAnd = requiredTasks.contains("&") || requiredTasks.contains(",");
        boolean hasOr = requiredTasks.contains("|");

        // If neither & nor | exist, treat it as |
        if (!hasAnd && !hasOr) {
            return new JSONObject().put("missing_some", new JSONArray(List.of(1, List.of(requiredTasks, "dummy")))).toString();
        }

        // If both & and | exist
        if (hasAnd && hasOr) {
            return generateJson(requiredTasks).toString();
        }

        // Only | or only &
        if (hasOr) {
            return new JSONObject().put("missing_some", new JSONArray(List.of(
                    1, Stream.concat(Arrays.stream(requiredTasks.split("\\|")), Stream.of("dummy")).toList()
            ))).toString();

        } else {
            return new JSONObject().put("missing", new JSONArray(List.of(requiredTasks.split("[,&]")))).toString();
        }
    }


    private JSONObject generateJson(String expression) {
        List<Object> conditions = parseExpression(expression);
        JSONObject json = new JSONObject();
        json.put("if", conditions);
        return json;
    }

    private List<Object> parseExpression(String expression) {
        List<Object> conditions = new ArrayList<>();
        expression = expression.replaceAll("\\s", "");

        List<String> missingValues = new ArrayList<>();
        List<String> missingSomeValues = new ArrayList<>();
        String[] splits;
        if (expression.contains("(")) {
            splits = expression.split("\\(");
            for (String s : splits) {
                if (s.contains(")") && (s.contains("&") || s.contains("|"))) {
                    s = s.replace("(", " ").replace(")", "").trim();
                    if (s.contains("&"))
                        missingValues.addAll(Arrays.asList(s.split("&")));

                    else if (s.contains("|")) {
                        missingSomeValues.addAll(Arrays.asList(s.split("\\|")));

                    }
                } else if (s.contains("&") && !s.contains(")")) {
                    missingValues.add(Arrays.stream(s.split("&")).findFirst().orElse(null));

                } else if (s.contains("|") && !s.contains(")")) {
                    missingSomeValues.add(Arrays.stream(s.split("\\|")).findFirst().orElse(null));

                }
            }
        } else {
            splits = expression.split("(?=[&|])|(?<=[&|])");
            if (expression.indexOf("&") < expression.indexOf("|")) {
                int index = IntStream.range(0, splits.length)
                        .filter(i -> "&".equals(splits[i]))
                        .findFirst()
                        .orElse(-1);
                missingValues.add(splits[index - 1]);
                missingValues.add(splits[index + 1]);
                missingSomeValues.add(splits[splits.length - 1]);

            } else {
                int index = IntStream.range(0, splits.length)
                        .filter(i -> "|".equals(splits[i]))
                        .findFirst()
                        .orElse(-1);
                missingSomeValues.add(splits[index - 1]);
                missingSomeValues.add(splits[index + 1]);
                missingValues.add(splits[splits.length - 1]);
            }


        }

        if (expression.indexOf("&") < expression.indexOf("|")) {
            if (!missingValues.isEmpty()) {
                conditions.add(new JSONObject().put("missing", new JSONArray(missingValues)));
            }
            if (!missingSomeValues.isEmpty()) {
                missingSomeValues.add("dummy");
                conditions.add(new JSONObject().put("missing_some", new JSONArray(List.of(1, missingSomeValues))));
            }
        } else {
            if (!missingSomeValues.isEmpty()) {
                missingSomeValues.add("dummy");
                conditions.add(new JSONObject().put("missing_some", new JSONArray(List.of(1, missingSomeValues))));
            }
            if (!missingValues.isEmpty()) {
                conditions.add(new JSONObject().put("missing", new JSONArray(missingValues)));
            }
        }

        conditions.add("OK");
        return conditions;
    }


}

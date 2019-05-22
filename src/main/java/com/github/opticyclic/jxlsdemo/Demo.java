package com.github.opticyclic.jxlsdemo;

import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;

import org.jxls.common.Context;
import org.jxls.util.JxlsHelper;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class Demo {
  private static Logger logger = LoggerFactory.getLogger(Demo.class);

  public static void main(String[] args) throws Exception {
    Path outputDir = Paths.get( "output");
    if (!Files.exists(outputDir)) {
      Files.createDirectory(outputDir);
      logger.info("Created directory " + outputDir.toAbsolutePath());
    }

    List<Employee> employees = generateSampleEmployeeData();

    //List
    logger.info("Running Collection Demo");
    try (InputStream is = Demo.class.getResourceAsStream("/object_collection_template.xlsx")) {
      try (OutputStream os = Files.newOutputStream(outputDir.resolve("object_collection_output.xlsx" ))) {
        Context context = new Context();
        context.putVar("employees", employees);
        JxlsHelper.getInstance().processTemplate(is, os, context);
      }
    }

    //Grouping
    logger.info("Running Grouping Demo");
    try (InputStream is = Demo.class.getResourceAsStream("/grouping_template.xlsx")) {
      try (OutputStream os = Files.newOutputStream(outputDir.resolve("grouping_output.xlsx" ))) {
        Context context = new Context();
        context.putVar("employees", employees);
        JxlsHelper.getInstance().processTemplate(is, os, context);
      }
    }
  }

  private static List<Employee> generateSampleEmployeeData() throws ParseException {
    List<Employee> employees = new ArrayList<>();
    SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MMM-dd", Locale.US);
    employees.add(new Employee("Elsa", dateFormat.parse("1970-Jul-10"), 1500, 0.15));
    employees.add(new Employee("Oleg", dateFormat.parse("1973-Apr-30"), 2300, 0.25));
    employees.add(new Employee("John", dateFormat.parse("1970-Jul-10"), 3500, 0.10));
    employees.add(new Employee("Neil", dateFormat.parse("1975-Oct-05"), 2500, 0.00));
    employees.add(new Employee("Maria", dateFormat.parse("1978-Jan-07"), 1700, 0.15));
    employees.add(new Employee("John", dateFormat.parse("1969-May-30"), 2800, 0.20));
    employees.add(new Employee("Oleg", dateFormat.parse("1988-Apr-30"), 1500, 0.15));
    employees.add(new Employee("Maria", dateFormat.parse("1970-Jul-10"), 3000, 0.10));
    employees.add(new Employee("John", dateFormat.parse("1973-Apr-30"), 1000, 0.05));
    return employees;
  }
}

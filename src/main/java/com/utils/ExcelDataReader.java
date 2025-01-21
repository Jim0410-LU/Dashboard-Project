package com.utils;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.read.listener.ReadListener;
import lombok.Data;
import lombok.Getter;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.time.DayOfWeek;

public class ExcelDataReader {
    // File path
    private String filePath;

    // Statistics for the day, using Map to store the number of tasks in each state
    private Map<String, Integer> currentStats;

    // Statistics by date, using Map to store statistics on different dates
    private Map<String, Map<String, Integer>> dailyStats;

    //Weekly statistics, using Map to store statistics for different weeks
    private Map<String, Map<String, Integer>> weeklyStats;

    // Task list, storing all task data
    private List<TaskData> taskList;

    // current date
    private String currentDate; // The current date in the format “d-MMM-yyy”

    // Constructor, accepts a file path parameter and initializes the statistics
    public ExcelDataReader(String filePath) {
        this.filePath = filePath;
        this.currentStats = new HashMap<>(); // Initialize the day's statistics
        this.dailyStats = new HashMap<>();
        this.weeklyStats = new HashMap<>();
        this.taskList = new ArrayList<>();
        this.currentDate = getCurrentWorkingDate();
        initializeStats();
    }

    // Get the current working date in “d-MMM-yyy” format
    private String getCurrentWorkingDate() {
        // Use LocalDate to get the current date and format it using the specified date formatting mode (“d-MMM-yyy”)
        return LocalDate.now().format(
                DateTimeFormatter.ofPattern("d-MMM-yy")
        );
    }

    // Initializes the day's statistics, setting the number of tasks in all states to 0 by default
    private void initializeStats() {

        currentStats.put("NEW", 0);
        currentStats.put("ONGOING", 0);
        currentStats.put("COMPLETED", 0);
        currentStats.put("WITHIN_TAT", 0);
        currentStats.put("OVER_TAT", 0);
    }

    // Reading and processing Excel data
    public void readExcelData() {
        // Use EasyExcel library to read Excel file with specified path, ExcelModel class to represent the mapping model of each row of data, TaskDataListener as data listener
        EasyExcel.read(filePath, ExcelModel.class, new TaskDataListener())
                .sheet()
                .doRead();

        calculatePercentages();
    }


    @Data
    public static class ExcelModel {


        @ExcelProperty("Date")
        private String date;


        @ExcelProperty("DocumentType")
        private String documentType;


        @ExcelProperty("ApplicationReceivedAt")
        private String applicationReceivedAt;


        @ExcelProperty("ScannedAt")
        private String scannedAt;


        @ExcelProperty("TotalTimeAtBranch")
        private String totalTimeAtBranch;


        @ExcelProperty("VerifiedAt")
        private String verifiedAt;


        @ExcelProperty("TotalTimeForVerification")
        private String totalTimeForVerification;


        @ExcelProperty("LodgementStartedAt")
        private String lodgementStartedAt;


        @ExcelProperty("ConfirmedAt")
        private String confirmedAt;


        @ExcelProperty("TotalTimeForEntry")
        private String totalTimeForEntry;


        @ExcelProperty("ComplianceVerifiedAt")
        private String complianceVerifiedAt;


        @ExcelProperty("AuthorizedAt")
        private String authorizedAt;


        @ExcelProperty("DocumentSerial")
        private String documentSerial;


        @ExcelProperty("Status")
        private String status;


        @ExcelProperty("ReferenceNumber")
        private String referenceNumber;


        @ExcelProperty("DetailType (Customer)")
        private String amount;


        @ExcelProperty("Description (ClientDetail)")
        private String clientName;


        @ExcelProperty("TAT")
        private String tat;


        @ExcelProperty("AuthorizedBy")
        private String handler;
    }


    @Data
    public static class TaskData {


        private String date;

        private String documentType;

        private String applicationReceivedAt;

        private String scannedAt;

        private String totalTimeAtBranch;

        private String verifiedAt;

        private String totalTimeForVerification;

        private String lodgementStartedAt;

        private String confirmedAt;

        private String totalTimeForEntry;

        private String complianceVerifiedAt;

        private String authorizedAt;

        private String documentSerial;

        private String referenceNumber;

        private String amount;

        private String clientName;

        private String status;

        private String tat;

        private String handler;

        // Constructor: Used to create the TaskData object and initialize all fields
        public TaskData(String documentSerial, String referenceNumber, String amount,
                        String clientName, String status, String tat, String handler,
                        String date, String documentType, String applicationReceivedAt,
                        String scannedAt, String totalTimeAtBranch, String verifiedAt,
                        String totalTimeForVerification, String lodgementStartedAt,
                        String confirmedAt, String totalTimeForEntry,
                        String complianceVerifiedAt, String authorizedAt) {
            this.documentSerial = documentSerial;
            this.referenceNumber = referenceNumber;
            this.amount = amount;
            this.clientName = clientName;
            this.status = status;
            this.tat = tat;
            this.handler = handler;
            this.date = date;
            this.documentType = documentType;
            this.applicationReceivedAt = applicationReceivedAt;
            this.scannedAt = scannedAt;
            this.totalTimeAtBranch = totalTimeAtBranch;
            this.verifiedAt = verifiedAt;
            this.totalTimeForVerification = totalTimeForVerification;
            this.lodgementStartedAt = lodgementStartedAt;
            this.confirmedAt = confirmedAt;
            this.totalTimeForEntry = totalTimeForEntry;
            this.complianceVerifiedAt = complianceVerifiedAt;
            this.authorizedAt = authorizedAt;
        }
    }


    private class TaskDataListener implements ReadListener<ExcelModel> {

        @Override
        public void invoke(ExcelModel data, AnalysisContext context) {
            // Processing task status and updating statistics based on Excel data
            processTaskStatus(data);

            // Convert the read Excel data into TaskData objects and add them to the task list
            taskList.add(new TaskData(
                    data.getDocumentSerial(),
                    data.getReferenceNumber(),
                    data.getAmount(),
                    data.getClientName(),
                    data.getStatus(),
                    data.getTat(),
                    data.getHandler(),
                    data.getDate(),
                    data.getDocumentType(),
                    data.getApplicationReceivedAt(),
                    data.getScannedAt(),
                    data.getTotalTimeAtBranch(),
                    data.getVerifiedAt(),
                    data.getTotalTimeForVerification(),
                    data.getLodgementStartedAt(),
                    data.getConfirmedAt(),
                    data.getTotalTimeForEntry(),
                    data.getComplianceVerifiedAt(),
                    data.getAuthorizedAt()
            ));
        }

        // This method is called after all data parsing is complete
        @Override
        public void doAfterAllAnalysed(AnalysisContext context) {
        }
    }

    // 处理任务状态的函数
    private void processTaskStatus(ExcelModel data) {
        String date = data.getDate();
        String status = data.getStatus().toUpperCase();
        String documentType = data.getDocumentType();

        // Processing of the day's task statistics
        if (date.equals(currentDate)) {

            currentStats.merge("NEW", 1, Integer::sum); // Add 1 to stats if it's a new mission
        }

        // Ongoing Tasks statistics
        if ("PENDING".equals(status)) {
            currentStats.merge("ONGOING", 1, Integer::sum);
        }

        // Completed Tasks 统计（当天）
        if ("LODGE".equals(status)) { // If the task status is “LODGE”
            currentStats.merge("COMPLETED", 1, Integer::sum); // Add 1 to the task completion statistic

            // TAT统计（当天）
            if ("Ecoll - Export Collection".equals(documentType)) { // If the document type is “Ecoll - Export Collection”
                // Determine if the task is within the TAT
                if (isWithinTargetTAT(data.getTat())) {
                    currentStats.merge("WITHIN_TAT", 1, Integer::sum); // Statistics plus 1 if within TAT
                } else {
                    currentStats.merge("OVER_TAT", 1, Integer::sum);
                }
            }
        }

        // Processing history statistics (by date)
        dailyStats.putIfAbsent(date, new HashMap<>()); // If the day's statistics do not exist, create a new HashMap
        Map<String, Integer> dayStats = dailyStats.get(date); // Get the day's statistics


        if ("PENDING".equals(status)) {
            dayStats.merge("ONGOING", 1, Integer::sum);
        }
        if ("LODGE".equals(status)) { // 如果任务状态是 "LODGE"（已完成）
            dayStats.merge("COMPLETED", 1, Integer::sum);
            if ("Ecoll - Export Collection".equals(documentType)) {

                if (isWithinTargetTAT(data.getTat())) {
                    dayStats.merge("WITHIN_TAT", 1, Integer::sum);
                } else {
                    dayStats.merge("OVER_TAT", 1, Integer::sum);
                }
            }
        }
    }

    // Methods for determining whether a task is new or not
    private boolean isNewTask(ExcelModel data) {
        return data.getDate() != null &&
                data.getDate().trim().equalsIgnoreCase(currentDate); // If the date of the task matches the current date, it is considered a new task
    }

    private void calculatePercentages() {
        // Calculation of total completed missions (within target TAT + exceeding target TAT)
        int totalLodged = currentStats.get("WITHIN_TAT") + currentStats.get("OVER_TAT");

        if (totalLodged > 0) {
            // Calculation of percentage of normal tasks
            int normalPercentage = (currentStats.get("WITHIN_TAT") * 100) / totalLodged;
            currentStats.put("NORMAL_PERCENTAGE", normalPercentage);
            currentStats.put("ABNORMAL_PERCENTAGE", 100 - normalPercentage);
        }
    }


    // Methods for checking that a task's TAT meets its objectives
    private void checkTAT(ExcelModel data) {
        try {
            // 1. Checking document type and status
            // If the status of the task is not “LODGE” or the document type is not “Ecoll - Export Collection”, it is returned directly without processing.
            if (!"LODGE".equals(data.getStatus().toUpperCase()) ||
                    !"Ecoll - Export Collection".equals(data.getDocumentType())) {
                return;
            }

            // 2. Get the TAT time and check if it is empty
            String tatString = data.getTat();
            if (tatString == null || tatString.trim().isEmpty()) {
                return;
            }

            // 3. Parsing TAT time string format “HH:mm:ss”
            String[] parts = tatString.split(":");
            if (parts.length >= 3) {  // Ensure that the TAT string contains hours, minutes, and seconds
                int hours = Integer.parseInt(parts[0]); // Hours
                int minutes = Integer.parseInt(parts[1]); // Minutes
                int seconds = Integer.parseInt(parts[2]); // Seconds

                // 4. Converting TAT to Total Seconds for Comparison
                int totalSeconds = hours * 3600 + minutes * 60 + seconds; // Calculate total seconds
                int targetSeconds = 4 * 3600; // Target TAT is 4 hours, i.e., 14,400 seconds.

                // 5. Determine if the task meets the goal based on the total TAT seconds
                if (totalSeconds <= targetSeconds) {
                    // If TAT is less than or equal to 4 hours, statistically normal (Within Target TAT)
                    currentStats.merge("NORMAL", 1, Integer::sum);
                } else {
                    // If TAT is greater than 4 hours, statistic is abnormal (Over Target TAT)
                    currentStats.merge("ABNORMAL", 1, Integer::sum);
                }
            }
        } catch (Exception e) {

            System.err.println("Error processing TAT for document: " + data.getDocumentSerial());
            e.printStackTrace();
        }
    }

// Getter method, which returns statistics for each type of task


    public int getNewTasksCount() {
        return currentStats.getOrDefault("NEW", 0); // 获取当前统计中的 "NEW" 任务数，如果没有则返回 0
    }


    public int getOngoingTasksCount() {
        return currentStats.getOrDefault("ONGOING", 0); // 获取当前统计中的 "ONGOING" 任务数，如果没有则返回 0
    }


    public int getCompletedTasksCount() {
        return currentStats.getOrDefault("COMPLETED", 0); // 获取当前统计中的 "COMPLETED" 任务数，如果没有则返回 0
    }


    public int getNormalTATCount() {
        return currentStats.getOrDefault("WITHIN_TAT", 0); // 获取当前统计中的 "WITHIN_TAT" 任务数，如果没有则返回 0
    }


    public int getAbnormalTATCount() {
        return currentStats.getOrDefault("OVER_TAT", 0); // 获取当前统计中的 "OVER_TAT" 任务数，如果没有则返回 0
    }


    public List<TaskData> getTaskList() {
        return taskList; // 返回任务列表
    }

    //  Get statistics for a specified date
    public Map<String, Integer> getDailyStats(String date) {
        return dailyStats.getOrDefault(date, new HashMap<>()); // Returns the statistics for the specified date, or an empty HashMap if none exists.
    }

    // Get statistics for a given week
    public Map<String, Integer> getWeeklyStats(String week) {
        return weeklyStats.getOrDefault(week, new HashMap<>());
    }


    private boolean isWithinTargetTAT(String tatString) {
        try {

            if (tatString == null || tatString.trim().isEmpty()) {
                return false;
            }


            String[] parts = tatString.split(":");
            if (parts.length >= 3) {
                int hours = Integer.parseInt(parts[0]);
                int minutes = Integer.parseInt(parts[1]);
                int seconds = Integer.parseInt(parts[2]);


                int totalSeconds = hours * 3600 + minutes * 60 + seconds;
                return totalSeconds <= 4 * 3600; // 4 小时 = 14400 秒
            }
        } catch (Exception e) {

            e.printStackTrace();
        }
        return false;
    }

    // Get all tasks for the specified date
    public List<TaskData> getTasksByDate(String date) {
        return taskList.stream()
                .filter(task -> date.equals(task.getDate())) // Filter out tasks matching the specified date
                .collect(Collectors.toList()); // Returns a list of eligible tasks
    }

    // Get all tasks for a given week
    public List<TaskData> getTasksByWeek(String week) {
        return taskList.stream()
                .filter(task -> week.equals(getWeekFromDate(task.getDate())))
                .collect(Collectors.toList());
    }


    // Get the week to which the date belongs
    private String getWeekFromDate(String dateStr) {
        try {
            // Define a date formatter to convert a date string in “d-MMM-yy” format to a LocalDate object.
            DateTimeFormatter formatter = DateTimeFormatter.ofPattern("d-MMM-yy");
            LocalDate date = LocalDate.parse(dateStr, formatter);

            // Get the first day of the month in which the date falls
            LocalDate firstDayOfMonth = date.withDayOfMonth(1);

            // Calculate what week of the month the date is (7 days per week, assuming January starts on day 1)
            int weekNumber = (date.getDayOfMonth() - 1) / 7 + 1;

            // Returns a string of the form “Week x”, where x is the week of the current date.
            return "Week " + weekNumber;
        } catch (Exception e) {
            e.printStackTrace();
            return "Week 1"; // If parsing fails, the first week is returned by default
        }
    }

    // Get the display label for the week based on the number of weeks, including start and end dates
    public String getWeekDisplayLabel(int weekNumber) {
        try {
            // Get the current date and calculate the first day of the current month
            LocalDate now = LocalDate.now();
            LocalDate firstDayOfMonth = now.withDayOfMonth(1);

            // Find the first business day (skip Saturday and Sunday)
            while (firstDayOfMonth.getDayOfWeek() == DayOfWeek.SATURDAY ||
                    firstDayOfMonth.getDayOfWeek() == DayOfWeek.SUNDAY) {
                firstDayOfMonth = firstDayOfMonth.plusDays(1);
            }

            // Calculate the start date of the week (from the first day, skipping the weekend) and calculate the end date (Friday)
            LocalDate weekStart = firstDayOfMonth.plusDays((weekNumber - 1) * 7);
            LocalDate weekEnd = weekStart.plusDays(4);


            DateTimeFormatter displayFormatter = DateTimeFormatter.ofPattern("MM.dd");


            return String.format("Week %d(%s-%s)",
                    weekNumber,
                    weekStart.format(displayFormatter),
                    weekEnd.format(displayFormatter)
            );
        } catch (Exception e) {
            e.printStackTrace();
            return "Week " + weekNumber;
        }
    }

    // Get all weekly labels for the current month
    public List<String> getMonthlyWeekLabels() {
        List<String> weekLabels = new ArrayList<>();
        try {
            // Get the current date, determine the first and last day of the current month
            LocalDate now = LocalDate.now();
            LocalDate firstDayOfMonth = now.withDayOfMonth(1);
            LocalDate lastDayOfMonth = now.withDayOfMonth(now.lengthOfMonth());


            LocalDate currentDate = firstDayOfMonth;
            int weekNumber = 1;


            while (currentDate.getMonth() == now.getMonth()) {

                if (currentDate.getDayOfWeek() != DayOfWeek.SATURDAY &&
                        currentDate.getDayOfWeek() != DayOfWeek.SUNDAY) {


                    LocalDate weekEnd = currentDate;
                    while (weekEnd.isBefore(lastDayOfMonth) &&
                            weekEnd.getDayOfWeek() != DayOfWeek.FRIDAY) {
                        weekEnd = weekEnd.plusDays(1);

                        if (weekEnd.getDayOfWeek() == DayOfWeek.SATURDAY) {
                            break;
                        }
                    }


                    DateTimeFormatter formatter = DateTimeFormatter.ofPattern("MM.dd");
                    weekLabels.add(String.format("Week %d(%s-%s)",
                            weekNumber++,
                            currentDate.format(formatter),
                            weekEnd.format(formatter)
                    ));


                    currentDate = weekEnd.plusDays(1);

                    while (currentDate.getDayOfWeek() == DayOfWeek.SATURDAY ||
                            currentDate.getDayOfWeek() == DayOfWeek.SUNDAY) {
                        currentDate = currentDate.plusDays(1);
                    }
                } else {
                    // 如果当前日期是周末，则跳到下一个工作日
                    currentDate = currentDate.plusDays(1);
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return weekLabels;
    }


    public List<TaskData> getTasksByDateRange(LocalDate startDate, LocalDate endDate) {

        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("d-MMM-yy");


        return taskList.stream()
                .filter(task -> {
                    try {

                        LocalDate taskDate = LocalDate.parse(task.getDate(), formatter);

                        return !taskDate.isBefore(startDate) && !taskDate.isAfter(endDate);
                    } catch (Exception e) {
                        return false;
                    }
                })
                .collect(Collectors.toList());
    }


    public List<WeekData> getMonthlyWeekData() {
        List<WeekData> weekDataList = new ArrayList<>();
        try {

            LocalDate now = LocalDate.now();


            LocalDate firstDayOfMonth = now.withDayOfMonth(1);
            LocalDate lastDayOfMonth = now.withDayOfMonth(now.lengthOfMonth());


            LocalDate currentDate = firstDayOfMonth;
            int weekNumber = 1;


            while (currentDate.getMonth() == now.getMonth()) {

                if (currentDate.getDayOfWeek() != DayOfWeek.SATURDAY &&
                        currentDate.getDayOfWeek() != DayOfWeek.SUNDAY) {


                    LocalDate weekEnd = currentDate;
                    while (weekEnd.isBefore(lastDayOfMonth) &&
                            weekEnd.getDayOfWeek() != DayOfWeek.FRIDAY) {
                        weekEnd = weekEnd.plusDays(1);


                        if (weekEnd.getDayOfWeek() == DayOfWeek.SATURDAY) {
                            break;
                        }
                    }


                    DateTimeFormatter formatter = DateTimeFormatter.ofPattern("MM.dd");
                    String weekLabel = String.format("Week %d(%s-%s)",
                            weekNumber++,
                            currentDate.format(formatter),
                            weekEnd.format(formatter)
                    );


                    List<TaskData> weekTasks = getTasksByDateRange(currentDate, weekEnd);


                    weekDataList.add(new WeekData(weekLabel, weekTasks));


                    currentDate = weekEnd.plusDays(1);

                    while (currentDate.getDayOfWeek() == DayOfWeek.SATURDAY ||
                            currentDate.getDayOfWeek() == DayOfWeek.SUNDAY) {
                        currentDate = currentDate.plusDays(1);
                    }
                } else {

                    currentDate = currentDate.plusDays(1);
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }

        return weekDataList;
    }


    public static class WeekData {
        private final String weekLabel;
        private final List<TaskData> tasks;


        public WeekData(String weekLabel, List<TaskData> tasks) {
            this.weekLabel = weekLabel;
            this.tasks = tasks;
        }

        public String getWeekLabel() { return weekLabel; }
        public List<TaskData> getTasks() { return tasks; }
    }

}
package com.ui;

import com.formdev.flatlaf.FlatLightLaf;
import com.utils.ExcelDataReader;
import org.jfree.chart.ChartFactory;
import org.jfree.chart.ChartPanel;
import org.jfree.chart.JFreeChart;
import org.jfree.chart.axis.CategoryLabelPositions;
import org.jfree.chart.axis.NumberAxis;
import org.jfree.chart.plot.CategoryPlot;
import org.jfree.chart.renderer.category.LineAndShapeRenderer;
import org.jfree.chart.renderer.category.BarRenderer;
import org.jfree.chart.renderer.category.StandardBarPainter;
import org.jfree.chart.ui.RectangleInsets;
import org.jfree.data.category.CategoryDataset;
import org.jfree.data.category.DefaultCategoryDataset;
import org.jfree.chart.labels.StandardCategoryItemLabelGenerator;
import org.jfree.chart.labels.StandardCategoryItemLabelGenerator;
import org.jfree.chart.labels.ItemLabelPosition;
import org.jfree.chart.labels.ItemLabelAnchor;
import org.jfree.ui.TextAnchor;
import java.awt.geom.Ellipse2D;
import java.awt.BasicStroke;
import java.awt.geom.Ellipse2D;
import java.awt.BasicStroke;
import java.awt.Font;

import javax.swing.*;
import java.awt.*;
import java.awt.event.MouseEvent;
import java.awt.event.MouseAdapter;
import javax.swing.border.AbstractBorder;
import javax.swing.table.DefaultTableModel;
import javax.swing.table.TableCellRenderer;
import javax.swing.table.TableColumn;

import java.awt.Graphics2D;
import java.awt.RenderingHints;
import java.util.Map;
import java.util.List;
import java.util.ArrayList;
import java.time.LocalDate;
import java.time.DayOfWeek;
import java.time.format.DateTimeFormatter;
import java.util.Collections;
import java.time.format.TextStyle;
import java.util.Locale;
import java.util.concurrent.Executors;
import java.util.concurrent.ScheduledExecutorService;
import java.util.concurrent.TimeUnit;

public class OptimizedDashboard extends JFrame {
    // Define theme colors
    private static final Color BACKGROUND_COLOR = Color.decode("#F8F9FA");  // Set the background color to light gray
    private static final Color CARD_BACKGROUND = Color.WHITE;  // Set the card background color to white
    private static final Color PRIMARY_COLOR = Color.decode("#2C3E50");  // Set the primary color to dark gray
    private static final Color SUCCESS_COLOR = Color.decode("#2ECC71");  // Set the color of the success status to green
    private static final Color WARNING_COLOR = Color.decode("#F1C40F");  // Set the color of the warning status to yellow
    private static final Color INFO_COLOR = Color.decode("#3498DB");  // Set the color of the message status to blue

    private static final Color NEW_TASK_COLOR = Color.decode("#4CAF50");  // Set the color of new tasks to a brighter green
    private static final Color ONGOING_TASK_COLOR = Color.decode("#FF9800");  // Set the color of ongoing tasks to a more vibrant orange
    private static final Color COMPLETED_TASK_COLOR = Color.decode("#2196F3");  // Set the color of completed tasks to a brighter blue

    private static final Color WITHIN_TAT_COLOR = Color.decode("#00C853");  // Setting the color in the TAT range to a more vibrant green
    private static final Color OVER_TAT_COLOR = Color.decode("#FF1744");  // Set colors outside the TAT range to a more vibrant red

    private boolean isWeeklyView = true;  // Controls whether the weekly view is displayed, defaults to weekly view
    private ChartPanel lineChartPanel;  // Storing panels for line graphs
    private ChartPanel barChartPanel;  // Panel for storing bar charts
    private JScrollPane tableScrollPane;  // Scrolling panels for storing data tables
    private JTable dataTable;  // Stored Data Forms
    private ExcelDataReader dataReader;  // Readers for reading Excel data
    private final ScheduledExecutorService scheduler = Executors.newScheduledThreadPool(1);  // Creating a Timed Task Scheduler
    private static final int REFRESH_INTERVAL = 10;  // Refresh interval (seconds)

    // Constructor, pass in the path to the Excel file and initialize the interface.
    public OptimizedDashboard(String excelFilePath) {
        try {
            UIManager.setLookAndFeel(new FlatLightLaf());  // Set the appearance to FlatLightLaf
        } catch (Exception ex) {
            ex.printStackTrace();  // If the setup fails, print an exception message
        }
        setTitle("Dashboard - Optimized Layout");  // Setting the window title
        setSize(1200, 900);  // Setting the window size
        setDefaultCloseOperation(EXIT_ON_CLOSE);  // Setting the window to exit the application when it closes
        setLayout(new GridBagLayout());  // Using the GridBagLayout Layout Manager
        getContentPane().setBackground(BACKGROUND_COLOR);  // Setting the background color of the window content area

        GridBagConstraints gbc = new GridBagConstraints();  // Creating the GridBagConstraints object
        gbc.fill = GridBagConstraints.BOTH;  // Letting components fill their available space
        gbc.insets = new Insets(5, 15, 5, 15);  // Set the inner margins of the component to reduce the top and bottom spacing and maintain the left and right spacing

        // Initializing the data reader and reading Excel data
        dataReader = new ExcelDataReader(excelFilePath);
        dataReader.readExcelData();

        // Part I: Statistical regions of the mandate
        JPanel taskSummaryPanel = createTaskSummaryPanel();  // Creating a Task Statistics Panel
        taskSummaryPanel.setPreferredSize(new Dimension(getWidth(), 90));
        gbc.gridx = 0;
        gbc.gridy = 0;
        gbc.gridwidth = 2;  // Set the component to span 2 columns
        gbc.weightx = 1;
        gbc.weighty = 0;
        add(taskSummaryPanel, gbc);  // Add the Task Statistics panel to the window

        // Part II: Mission Status Indication Area
        JPanel taskStatusPanel = createTaskStatusPanel();  // Creating a Task Status Panel
        taskStatusPanel.setPreferredSize(new Dimension(getWidth(), 90));  // Set the preferred size of the panel, reduce the height
        gbc.gridx = 0;
        gbc.gridy = 1;
        gbc.gridwidth = 2;
        gbc.weightx = 1;
        gbc.weighty = 0;
        add(taskStatusPanel, gbc);  // Add the Task Statistics panel to the window

        // Part III: Toggle Button Area
        JPanel togglePanel = createTogglePanel();  // Creating Toggle Button Panels
        togglePanel.setPreferredSize(new Dimension(getWidth(), 35));
        gbc.gridx = 0;
        gbc.gridy = 2;
        gbc.gridwidth = 2;
        gbc.weightx = 1;
        gbc.weighty = 0;
        add(togglePanel, gbc);  // Adding the toggle button panel to a window

        // Part IV: Line Charts
        lineChartPanel = createLineChart(350);  // Creating and setting up the Line Chart panel
        gbc.gridx = 0;
        gbc.gridy = 3;
        gbc.gridwidth = 2;
        gbc.weightx = 1;
        gbc.weighty = 1;
        gbc.insets = new Insets(5, 15, 5, 15);  // Set the inner margins of line charts to reduce the spacing between charts
        add(lineChartPanel, gbc);

        // Part V: Bar charts
        barChartPanel = createBarChart(350);  // Creating and setting up the Bar chart panel
        gbc.gridx = 0;
        gbc.gridy = 4;
        gbc.gridwidth = 2;
        gbc.weightx = 1;
        gbc.weighty = 1;
        add(barChartPanel, gbc);

        // Part VI: Data table areas
        tableScrollPane = new JScrollPane(createDataTable());  // Create a table of data and place it in a scrolling panel
        tableScrollPane.setPreferredSize(new Dimension(getWidth(), 100));
        gbc.gridx = 0;
        gbc.gridy = 5;
        gbc.gridwidth = 2;
        gbc.weightx = 1;
        gbc.weighty = 0.5;
        gbc.insets = new Insets(5, 15, 5, 15);
        add(tableScrollPane, gbc);

        // Add a timed refresh task
        startAutoRefresh();  // Activate the timed refresh function
    }


    // Method to start a timed dashboard refresh
    private void startAutoRefresh() {
        // Use the scheduler to execute tasks periodically at intervals specified by REFRESH_INTERVAL
        scheduler.scheduleAtFixedRate(() -> {
            try {
                String excelFilePath = "data.xlsx";  // Specify the path to the Excel file

                // Perform UI updates in the Event Dispatch Thread (EDT) to ensure thread-safety
                SwingUtilities.invokeLater(() -> {
                    // Reinitialize the data reader to read new Excel data
                    dataReader = new ExcelDataReader(excelFilePath);
                    dataReader.readExcelData();  // Reading data from an Excel file

                    // Updating the Dashboard
                    updateDashboard();
                });
            } catch (Exception e) {
                e.printStackTrace();
            }
        }, REFRESH_INTERVAL, REFRESH_INTERVAL, TimeUnit.SECONDS);  // Scheduling tasks at fixed intervals
    }


    // Add a way to update the dashboard
    // Ways to update dashboard
    private void updateDashboard() {
        try {
            // Update the data in the Task Statistics panel
            updateTaskSummaryCards();

            // Update data in the task status panel
            updateTaskStatusCards();


            updateCharts();

            // Updating data tables
            updateDataTable();

            // Refresh UI Interface
            revalidate();
            repaint();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }


    // Method of updating the task statistics card
    private void updateTaskSummaryCards() {
        // Get the Task Statistics panelGet the Task Statistics panel
        JPanel taskSummaryPanel = (JPanel) ((JPanel)getContentPane().getComponent(0));
        Component[] cards = taskSummaryPanel.getComponents();  // Get all card components

        // Iterate through each card and update the values
        for (Component card : cards) {
            if (card instanceof JPanel) {
                JPanel taskCard = (JPanel) card;
                JLabel titleLabel = (JLabel) taskCard.getComponent(0);
                JLabel countLabel = (JLabel) taskCard.getComponent(1);


                switch (titleLabel.getText()) {
                    case "New Tasks":
                        countLabel.setText(String.valueOf(dataReader.getNewTasksCount()));
                        break;
                    case "Ongoing Tasks":
                        countLabel.setText(String.valueOf(dataReader.getOngoingTasksCount()));
                        break;
                    case "Completed Tasks":
                        countLabel.setText(String.valueOf(dataReader.getCompletedTasksCount()));
                        break;
                }
            }
        }
    }


    // Methods for updating task status cards
    private void updateTaskStatusCards() {

        JPanel taskStatusPanel = (JPanel) ((JPanel)getContentPane().getComponent(1));
        Component[] cards = taskStatusPanel.getComponents();


        for (Component card : cards) {
            if (card instanceof JPanel) {
                JPanel statusCard = (JPanel) card;
                JLabel titleLabel = (JLabel) statusCard.getComponent(0);  // 获取卡片标题标签
                JLabel countLabel = (JLabel) statusCard.getComponent(1);  // 获取卡片数值标签


                switch (titleLabel.getText()) {
                    case "Within Target TAT":
                        countLabel.setText(String.valueOf(dataReader.getNormalTATCount()));
                        break;
                    case "Over Target TAT":
                        countLabel.setText(String.valueOf(dataReader.getAbnormalTATCount()));
                        break;
                }
            }
        }
    }


    // Methods for updating data tables
    private void updateDataTable() {

        // Getting task list data from the data reader
        List<ExcelDataReader.TaskData> taskList = dataReader.getTaskList();
        Object[][] newData = new Object[taskList.size()][dataTable.getColumnCount()];

        // Populate the task data into the array
        for (int i = 0; i < taskList.size(); i++) {
            ExcelDataReader.TaskData task = taskList.get(i);
            newData[i] = new Object[] {
                    convertDateToEnglishFormat(task.getDate()),
                    task.getDocumentSerial(),
                    task.getDocumentType(),
                    task.getReferenceNumber(),
                    task.getAmount(),
                    task.getClientName(),
                    task.getStatus(),
                    task.getTat(),
                    task.getHandler(),
                    task.getApplicationReceivedAt(),
                    task.getScannedAt(),
                    task.getTotalTimeAtBranch(),
                    task.getVerifiedAt(),
                    task.getTotalTimeForVerification(),
                    task.getLodgementStartedAt(),
                    task.getConfirmedAt(),
                    task.getTotalTimeForEntry(),
                    task.getComplianceVerifiedAt(),
                    task.getAuthorizedAt()
            };
        }


        String[] columnNames = {
                "Date", "Document Serial", "Document Type", "Reference Number",
                "Amount", "Client Name", "Status", "TAT", "Handler",
                "Application Received At", "Scanned At", "Total Time At Branch",
                "Verified At", "Total Time For Verification",
                "Lodgement Started At", "Confirmed At", "Total Time For Entry",
                "Compliance Verified At", "Authorized At"
        };

        // Updating the table model with new data and column names
        dataTable.setModel(new DefaultTableModel(newData, columnNames));

        // Setting the basic properties of the form
        dataTable.setFillsViewportHeight(true);
        dataTable.setRowHeight(35);
        dataTable.setFont(new Font("Segoe UI", Font.PLAIN, 14));
        dataTable.getTableHeader().setFont(new Font("Segoe UI", Font.BOLD, 14));
        dataTable.setShowGrid(false);
        dataTable.setIntercellSpacing(new Dimension(0, 0));
        dataTable.setSelectionBackground(new Color(52, 152, 219, 50));
        dataTable.setSelectionForeground(PRIMARY_COLOR);


        dataTable.getTableHeader().setBackground(BACKGROUND_COLOR);
        dataTable.getTableHeader().setForeground(PRIMARY_COLOR);
        dataTable.getTableHeader().setBorder(BorderFactory.createEmptyBorder());

        // Setting Column Width
        dataTable.setAutoResizeMode(JTable.AUTO_RESIZE_OFF);
        for (int i = 0; i < dataTable.getColumnCount(); i++) {
            dataTable.getColumnModel().getColumn(i).setPreferredWidth(150);
        }
    }


    // Modify window close handling
    @Override
    public void dispose() {
        // Turn off the timer and stop scheduling tasks
        scheduler.shutdown();
        try {
            // Wait for the timer to turn off completely, up to 2 seconds
            if (!scheduler.awaitTermination(2, TimeUnit.SECONDS)) {
                // If the timer is not closed within 2 seconds, force the timer to be closed
                scheduler.shutdownNow();
            }
        } catch (InterruptedException e) {
            // Force the timer to close if the thread is interrupted while waiting
            scheduler.shutdownNow();
        }

        super.dispose();
    }


    // Creating panels for the Task Statistics area
    private JPanel createTaskSummaryPanel() {
        // Use the GridLayout layout manager to create a 1-row, 3-column panel
        JPanel panel = new JPanel(new GridLayout(1, 3, 15, 0));

        panel.add(createTaskCard("New Tasks", dataReader.getNewTasksCount(), NEW_TASK_COLOR));  // 新任务卡片
        panel.add(createTaskCard("Ongoing Tasks", dataReader.getOngoingTasksCount(), ONGOING_TASK_COLOR));  // 进行中任务卡片
        panel.add(createTaskCard("Completed Tasks", dataReader.getCompletedTasksCount(), COMPLETED_TASK_COLOR));  // 已完成任务卡片


        return panel;
    }


    // Create panels for task status indication areas
    private JPanel createTaskStatusPanel() {

        JPanel panel = new JPanel(new GridLayout(1, 2, 15, 0));
        panel.add(createStatusCard("Within Target TAT", dataReader.getNormalTATCount(), WITHIN_TAT_COLOR));  // 在目标时限内任务卡片
        panel.add(createStatusCard("Over Target TAT", dataReader.getAbnormalTATCount(), OVER_TAT_COLOR));  // 超出目标时限任务卡片

        return panel;
    }


    // Create toggle button area
    private JPanel createTogglePanel() {

        JPanel panel = new JPanel();
        panel.setBackground(BACKGROUND_COLOR);
        panel.setBorder(BorderFactory.createEmptyBorder(2, 0, 2, 0));

        // Create two buttons: one for the weekly view and the other for the monthly view
        JButton weeklyButton = createStyledButton("Weekly");
        JButton monthlyButton = createStyledButton("Monthly");

        // Adding a click event listener for the Week View button
        weeklyButton.addActionListener(e -> {
            isWeeklyView = true;  // Set to weekly view
            updateCharts();
        });

        // Adding a click event listener for the Month View button
        monthlyButton.addActionListener(e -> {
            isWeeklyView = false;
            updateCharts();
        });

        // Setting the preferred size of the button
        Dimension buttonSize = new Dimension(100, 28);
        weeklyButton.setPreferredSize(buttonSize);
        monthlyButton.setPreferredSize(buttonSize);
        panel.add(weeklyButton);
        panel.add(Box.createHorizontalStrut(8));
        panel.add(monthlyButton);


        return panel;
    }


    // Update chart area
    private void updateCharts() {
        // Remove existing Line Chart and Bar Chart panels
        remove(lineChartPanel);
        remove(barChartPanel);

        // Create new Line Chart and Bar Chart panels, passing in 350 as the height
        lineChartPanel = createLineChart(350);
        barChartPanel = createBarChart(350);

        // Set GridBagConstraints for adjusting the layout of the component
        GridBagConstraints gbc = new GridBagConstraints();
        gbc.fill = GridBagConstraints.BOTH;
        gbc.insets = new Insets(10, 10, 10, 10);

        // Re-add the Line Chart panel to the layout
        gbc.gridx = 0;
        gbc.gridy = 3;
        gbc.gridwidth = 2;
        gbc.weightx = 1;
        gbc.weighty = 1;
        add(lineChartPanel, gbc);

        // Re-add the Bar Chart panel to the layout
        gbc.gridx = 0;
        gbc.gridy = 4;
        gbc.gridwidth = 2;
        gbc.weightx = 1;
        gbc.weighty = 1;
        add(barChartPanel, gbc);

        // Re-validating the layout and redrawing the interface
        revalidate();
        repaint();
    }


    // Creating Task Cards
    private JPanel createTaskCard(String title, int count, Color color) {

        JPanel card = new JPanel(new BorderLayout(5, 5));
        card.setBackground(CARD_BACKGROUND);

        // Setting the border of the card: adding rounded corners and shadow effects
        card.setBorder(BorderFactory.createCompoundBorder(
                BorderFactory.createLineBorder(new Color(0, 0, 0, 20), 1),
                BorderFactory.createEmptyBorder(10, 10, 10, 10)
        ));

        // Create title tags and set fonts and colors
        JLabel titleLabel = new JLabel(title);
        titleLabel.setFont(new Font("Segoe UI", Font.BOLD, 16));
        titleLabel.setForeground(PRIMARY_COLOR);

        // Create count labels and set fonts and colors
        JLabel countLabel = new JLabel(String.valueOf(count));
        countLabel.setFont(new Font("Segoe UI", Font.BOLD, 32));
        countLabel.setForeground(color);

        card.add(titleLabel, BorderLayout.NORTH);
        card.add(countLabel, BorderLayout.CENTER);


        return card;
    }


    // Creating Status Card
    private JPanel createStatusCard(String title, int count, Color color) {
        // Create a panel, using the BorderLayout layout, and set the vertical spacing to 5 pixels
        JPanel card = new JPanel(new BorderLayout(0, 5));
        card.setBackground(color);
        card.setBorder(BorderFactory.createCompoundBorder(
                new ShadowBorder(),
                BorderFactory.createEmptyBorder(10, 10, 10, 10)
        ));


        JLabel titleLabel = new JLabel(title, SwingConstants.CENTER);
        titleLabel.setFont(new Font("Segoe UI", Font.BOLD, 16));
        titleLabel.setForeground(Color.WHITE);
        card.add(titleLabel, BorderLayout.NORTH);

        // Create count labels, center them, and remove the percentage sign
        JLabel countLabel = new JLabel(String.valueOf(count), SwingConstants.CENTER);
        countLabel.setFont(new Font("Segoe UI", Font.BOLD, 32));
        countLabel.setForeground(Color.WHITE);
        card.add(countLabel, BorderLayout.CENTER);

        // Add mouse hover effect
        card.addMouseListener(new MouseAdapter() {
            public void mouseEntered(MouseEvent e) {
                card.setBackground(color.darker());
            }
            public void mouseExited(MouseEvent e) {
                card.setBackground(color);
            }
        });

        return card;
    }


    // Creating Line Charts
    private ChartPanel createLineChart(int height) {

        JFreeChart lineChart = ChartFactory.createLineChart(
                isWeeklyView ? "Weekly Completed Tasks Trend" : "Monthly Completed Tasks Trend",
                isWeeklyView ? "Day of Week" : "Week",
                "Count",
                createLineDataset()
        );


        CategoryPlot plot = lineChart.getCategoryPlot();
        plot.setBackgroundPaint(Color.WHITE);
        plot.setDomainGridlinePaint(new Color(0, 0, 0, 20));
        plot.setRangeGridlinePaint(new Color(0, 0, 0, 20));


        NumberAxis rangeAxis = (NumberAxis) plot.getRangeAxis();
        rangeAxis.setUpperMargin(0.20);

        // Setting the Chart Font
        lineChart.getTitle().setFont(new Font("Segoe UI", Font.BOLD, 16));
        plot.getDomainAxis().setLabelFont(new Font("Segoe UI", Font.PLAIN, 12));
        plot.getRangeAxis().setLabelFont(new Font("Segoe UI", Font.PLAIN, 12));

        // Get the renderer (for setting line styles)
        LineAndShapeRenderer renderer = (LineAndShapeRenderer) plot.getRenderer();

        // Setting line colors and styles
        renderer.setSeriesPaint(0, COMPLETED_TASK_COLOR);
        renderer.setSeriesStroke(0, new BasicStroke(2.0f));
        renderer.setSeriesShape(0, new Ellipse2D.Double(-3.0, -3.0, 6.0, 6.0));

        // Setting up data labels
        renderer.setDefaultItemLabelGenerator(new StandardCategoryItemLabelGenerator());
        renderer.setDefaultItemLabelsVisible(true);
        renderer.setDefaultItemLabelFont(new Font("Segoe UI", Font.PLAIN, 11));

        // Set the position of the label (normal display, vertical display)
        renderer.setDefaultPositiveItemLabelPosition(
                new ItemLabelPosition(
                        ItemLabelAnchor.OUTSIDE12,
                        org.jfree.chart.ui.TextAnchor.BOTTOM_CENTER
                )
        );

        // Setting the Margins of a Chart
        plot.setInsets(new RectangleInsets(10, 10, 10, 10));

        // Creating and setting up chart panels
        ChartPanel chartPanel = new ChartPanel(lineChart);
        chartPanel.setPreferredSize(new Dimension(getWidth(), height));
        chartPanel.setBorder(BorderFactory.createEmptyBorder(5, 5, 5, 5));
        chartPanel.setBackground(CARD_BACKGROUND);


        return chartPanel;
    }


    // Creating Line Chart Data Sets
    private CategoryDataset createLineDataset() {
        DefaultCategoryDataset dataset = new DefaultCategoryDataset();  // Create an empty dataset

        if (isWeeklyView) {
            // Get the last 5 working days and their corresponding display labels
            List<WorkDay> workDays = getLastFiveWorkDays();

            // Adding data in chronological order
            for (WorkDay workDay : workDays) {
                String date = workDay.getDate();
                String label = workDay.getLabel() + "（"+ convertDateToEnglishFormat(date) +"）";
                List<ExcelDataReader.TaskData> dayTasks = dataReader.getTasksByDate(date);

                // Count the number of tasks with a status of LODGE
                long completedTasks = dayTasks.stream()
                        .filter(task -> "LODGE".equalsIgnoreCase(task.getStatus()))
                        .count();

                dataset.addValue(completedTasks, "Completed Tasks", label);
            }
        } else {
            // Monthly view: using the new method of data acquisition
            List<ExcelDataReader.WeekData> weekDataList = dataReader.getMonthlyWeekData();  // Access to weekly data
            for (ExcelDataReader.WeekData weekData : weekDataList) {
                String weekLabel = weekData.getWeekLabel();
                List<ExcelDataReader.TaskData> weekTasks = weekData.getTasks();

                // Counts the number of tasks with statuses of PENDING and LODGE
                long ongoingTasks = weekTasks.stream()
                        .filter(task -> "PENDING".equalsIgnoreCase(task.getStatus()))
                        .count();

                long completedTasks = weekTasks.stream()
                        .filter(task -> "LODGE".equalsIgnoreCase(task.getStatus()))
                        .count();

                // Statistics on TAT-compliant and non-TAT-compliant tasks
                long withinTAT = weekTasks.stream()
                        .filter(task -> "LODGE".equalsIgnoreCase(task.getStatus())
                                && "Ecoll - Export Collection".equals(task.getDocumentType())
                                && isWithinTargetTAT(task.getTat()))
                        .count();

                long overTAT = weekTasks.stream()
                        .filter(task -> "LODGE".equalsIgnoreCase(task.getStatus())
                                && "Ecoll - Export Collection".equals(task.getDocumentType())
                                && !isWithinTargetTAT(task.getTat()))
                        .count();


                dataset.addValue(completedTasks, "Completed Tasks", weekLabel);
            }
        }

        return dataset;
    }


    // Creating Bar Charts
    private ChartPanel createBarChart(int height) {

        JFreeChart barChart = ChartFactory.createBarChart(
                "Last 5 Working Days Status Distribution",
                "Date",
                "Count",
                createBarDataset()
        );

        CategoryPlot plot = barChart.getCategoryPlot();
        plot.setBackgroundPaint(Color.WHITE);
        plot.setDomainGridlinePaint(new Color(0, 0, 0, 20));
        plot.setRangeGridlinePaint(new Color(0, 0, 0, 20));

        NumberAxis rangeAxis = (NumberAxis) plot.getRangeAxis();
        rangeAxis.setUpperMargin(0.20);

        barChart.getTitle().setFont(new Font("Segoe UI", Font.BOLD, 16));
        plot.getDomainAxis().setLabelFont(new Font("Segoe UI", Font.PLAIN, 12));
        plot.getRangeAxis().setLabelFont(new Font("Segoe UI", Font.PLAIN, 12));

        plot.getDomainAxis().setCategoryLabelPositions(CategoryLabelPositions.STANDARD);


        plot.setInsets(new RectangleInsets(5, 5, 5, 5));


        BarRenderer renderer = (BarRenderer) plot.getRenderer();

        renderer.setSeriesPaint(0, NEW_TASK_COLOR);
        renderer.setSeriesPaint(1, ONGOING_TASK_COLOR);
        renderer.setSeriesPaint(2, COMPLETED_TASK_COLOR);
        renderer.setSeriesPaint(3, WITHIN_TAT_COLOR);
        renderer.setSeriesPaint(4, OVER_TAT_COLOR);

        // Setting the numeric label display
        renderer.setDefaultItemLabelGenerator(new StandardCategoryItemLabelGenerator());
        renderer.setDefaultItemLabelsVisible(true);
        renderer.setDefaultItemLabelFont(new Font("Segoe UI", Font.PLAIN, 11));

        // Adjust the position of the value labels
        renderer.setDefaultPositiveItemLabelPosition(
                new ItemLabelPosition(
                        ItemLabelAnchor.OUTSIDE12,
                        org.jfree.chart.ui.TextAnchor.BOTTOM_CENTER
                )
        );


        renderer.setBarPainter(new StandardBarPainter());
        renderer.setShadowVisible(false);
        renderer.setMaximumBarWidth(0.1);

        // Create a ChartPanel, set its size and background color.
        ChartPanel chartPanel = new ChartPanel(barChart);
        chartPanel.setPreferredSize(new Dimension(getWidth(), height));
        chartPanel.setBorder(BorderFactory.createEmptyBorder(5, 5, 5, 5));
        chartPanel.setBackground(CARD_BACKGROUND);

        return chartPanel;
    }

    // Modifying Bar Chart Data Sets
    private CategoryDataset createBarDataset() {
        // Create a default dataset
        DefaultCategoryDataset dataset = new DefaultCategoryDataset();

        if (isWeeklyView) {
            // Get the last 5 working days and their corresponding display labels
            List<WorkDay> workDays = getLastFiveWorkDays();

            for (WorkDay workDay : workDays) {
                String date = workDay.getDate();
                String label = workDay.getLabel() + "（" + convertDateToEnglishFormat(date) + "）";
                List<ExcelDataReader.TaskData> dayTasks = dataReader.getTasksByDate(date);

                long newTasks = dayTasks.stream()
                        .filter(task -> task.getDate().equals(date)) // Filter out the day's tasks
                        .count();

                long ongoingTasks = dayTasks.stream()
                        .filter(task -> "PENDING".equalsIgnoreCase(task.getStatus()))
                        .count();


                long completedTasks = dayTasks.stream()
                        .filter(task -> "LODGE".equalsIgnoreCase(task.getStatus()))
                        .count();

                // 4&5. Within/Over Target TAT:
                // DocumentType == Ecoll - Export Collection && Status == LODGE
                long withinTAT = dayTasks.stream()
                        .filter(task -> "LODGE".equalsIgnoreCase(task.getStatus())
                                && "Ecoll - Export Collection".equals(task.getDocumentType()) // Ecoll - Export Collection
                                && isWithinTargetTAT(task.getTat())) //
                        .count();

                long overTAT = dayTasks.stream()
                        .filter(task -> "LODGE".equalsIgnoreCase(task.getStatus())
                                && "Ecoll - Export Collection".equals(task.getDocumentType())
                                && !isWithinTargetTAT(task.getTat()))
                        .count();

                // Adding data to a dataset
                dataset.addValue(newTasks, "New Tasks", label);
                dataset.addValue(ongoingTasks, "Ongoing Tasks", label);
                dataset.addValue(completedTasks, "Completed Tasks", label);
                dataset.addValue(withinTAT, "Within Target TAT", label);
                dataset.addValue(overTAT, "Over Target TAT", label);
            }
        } else {

            List<ExcelDataReader.WeekData> weekDataList = dataReader.getMonthlyWeekData();
            for (ExcelDataReader.WeekData weekData : weekDataList) {
                String weekLabel = weekData.getWeekLabel();
                List<ExcelDataReader.TaskData> weekTasks = weekData.getTasks();


                long ongoingTasks = weekTasks.stream()
                        .filter(task -> "PENDING".equalsIgnoreCase(task.getStatus()))
                        .count();

                long completedTasks = weekTasks.stream()
                        .filter(task -> "LODGE".equalsIgnoreCase(task.getStatus()))
                        .count();

                long withinTAT = weekTasks.stream()
                        .filter(task -> "LODGE".equalsIgnoreCase(task.getStatus())
                                && "Ecoll - Export Collection".equals(task.getDocumentType())
                                && isWithinTargetTAT(task.getTat()))
                        .count();

                long overTAT = weekTasks.stream()
                        .filter(task -> "LODGE".equalsIgnoreCase(task.getStatus())
                                && "Ecoll - Export Collection".equals(task.getDocumentType())
                                && !isWithinTargetTAT(task.getTat()))
                        .count();


                dataset.addValue(ongoingTasks, "Ongoing Tasks", weekLabel);
                dataset.addValue(completedTasks, "Completed Tasks", weekLabel);
                dataset.addValue(withinTAT, "Within Target TAT", weekLabel);
                dataset.addValue(overTAT, "Over Target TAT", weekLabel);
            }
        }

        return dataset;
    }


    // Auxiliary method: determining whether the TAT is within the target time
    private boolean isWithinTargetTAT(String tatString) {
        try {
            // Returns false if the TAT string is null or a space
            if (tatString == null || tatString.trim().isEmpty()) {
                return false;
            }

            // Split the TAT string by colons to get hours, minutes and seconds
            String[] parts = tatString.split(":");
            if (parts.length >= 3) {

                int hours = Integer.parseInt(parts[0]);
                int minutes = Integer.parseInt(parts[1]);
                int seconds = Integer.parseInt(parts[2]);

                // Calculate total TAT seconds
                int totalSeconds = hours * 3600 + minutes * 60 + seconds;

                // Returns true if the total number of TAT seconds is less than or equal to 4 hours (14,400 seconds), indicating that the target time in the
                return totalSeconds <= 4 * 3600;
            }
        } catch (Exception e) {
            e.printStackTrace(); // Catch exceptions and print stack information
        }
        return false;
    }


    private static class WorkDay {
        private final String date;
        private final String label;

        public WorkDay(String date, String label) {
            this.date = date;
            this.label = label;
        }

        // Get the actual date
        public String getDate() { return date; }

        // Get display labels
        public String getLabel() { return label; }
    }

    // Get the last 5 business days and their labels
    private List<WorkDay> getLastFiveWorkDays() {
        List<WorkDay> workDays = new ArrayList<>();
        LocalDate currentDate = LocalDate.now();
        DateTimeFormatter dateFormatter = DateTimeFormatter.ofPattern("d-MMM-yy");


        while (workDays.size() < 5) {
            //  If the current date is not a Saturday or Sunday, it is considered a weekday
            if (currentDate.getDayOfWeek() != DayOfWeek.SATURDAY &&
                    currentDate.getDayOfWeek() != DayOfWeek.SUNDAY) {

                String date = currentDate.format(dateFormatter);

                String label = currentDate.getDayOfWeek()
                        .getDisplayName(TextStyle.SHORT, Locale.ENGLISH);

                workDays.add(0, new WorkDay(date, label));
            }
            // Current date minus one day to check if the previous day was a business day
            currentDate = currentDate.minusDays(1);
        }

        return workDays;
    }


    private String convertChineseMonthToEnglish(String chineseMonth) {
        switch (chineseMonth) {
            case "一月": return "Jan";
            case "二月": return "Feb";
            case "三月": return "Mar";
            case "四月": return "Apr";
            case "五月": return "May";
            case "六月": return "Jun";
            case "七月": return "Jul";
            case "八月": return "Aug";
            case "九月": return "Sep";
            case "十月": return "Oct";
            case "十一月": return "Nov";
            case "十二月": return "Dec";
            default: return chineseMonth;
        }
    }

    public String convertDateToEnglishFormat(String dateString) {

        String[] parts = dateString.split("-");
        String day = parts[0];
        String month = convertChineseMonthToEnglish(parts[1]);
        String year = parts[2];


        return day + "-" + month + "-" + year;
    }



    private JTable createDataTable() {

        String[] columnNames = {
                "Date", "Document Serial", "Document Type", "Reference Number", "Amount", "Client Name", "Status",
                "TAT", "Handler", "Application Received At", "Scanned At", "Total Time At Branch", "Verified At",
                "Total Time For Verification", "Lodgement Started At", "Confirmed At", "Total Time For Entry",
                "Compliance Verified At", "Authorized At"
        };


        List<ExcelDataReader.TaskData> taskList = dataReader.getTaskList();
        Object[][] data = new Object[taskList.size()][columnNames.length];


        for (int i = 0; i < taskList.size(); i++) {
            ExcelDataReader.TaskData task = taskList.get(i);
            data[i] = new Object[] {
                    convertDateToEnglishFormat(task.getDate()),  // 日期转换为英文格式
                    task.getDocumentSerial(),
                    task.getDocumentType(),
                    task.getReferenceNumber(),
                    task.getAmount(),
                    task.getClientName(),
                    task.getStatus(),
                    task.getTat(),
                    task.getHandler(),
                    task.getApplicationReceivedAt(),
                    task.getScannedAt(),
                    task.getTotalTimeAtBranch(),
                    task.getVerifiedAt(),
                    task.getTotalTimeForVerification(),
                    task.getLodgementStartedAt(),
                    task.getConfirmedAt(),
                    task.getTotalTimeForEntry(),
                    task.getComplianceVerifiedAt(),
                    task.getAuthorizedAt()
            };
        }

        // Create a custom JTable for displaying data
        dataTable = new JTable(data, columnNames) {
            @Override
            public Component prepareRenderer(TableCellRenderer renderer, int row, int column) {
                Component c = super.prepareRenderer(renderer, row, column);

                // Determine if the current column is a TAT column (column 8)
                if (column == 7) {
                    String documentType = (String) getValueAt(row, 2);
                    String status = (String) getValueAt(row, 6);
                    String tatValue = (String) getValueAt(row, 7);

                    // Only records in the 'Ecoll - Export Collection' with status 'LODGE' are checked.
                    if ("Ecoll - Export Collection".equals(documentType) && "LODGE".equalsIgnoreCase(status)) {
                        // If the target TAT time is exceeded, set the background color to light red
                        if (!isWithinTargetTAT(tatValue)) {
                            c.setBackground(new Color(255, 204, 204));
                        } else {
                            c.setBackground(getBackground());
                        }
                    } else {
                        c.setBackground(getBackground());
                    }
                } else {
                    c.setBackground(getBackground());
                }

                return c;
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
                        return totalSeconds <= 4 * 3600;
                    }
                } catch (Exception e) {
                    System.err.println("Error parsing TAT: " + tatString);
                }
                return false;
            }
        };


        dataTable.setFillsViewportHeight(true);
        dataTable.setRowHeight(35);
        dataTable.setFont(new Font("Segoe UI", Font.PLAIN, 14));
        dataTable.getTableHeader().setFont(new Font("Segoe UI", Font.BOLD, 14));
        dataTable.setShowGrid(false);
        dataTable.setIntercellSpacing(new Dimension(0, 0));
        dataTable.setSelectionBackground(new Color(52, 152, 219, 50));
        dataTable.setSelectionForeground(PRIMARY_COLOR);


        dataTable.getTableHeader().setBackground(BACKGROUND_COLOR);
        dataTable.getTableHeader().setForeground(PRIMARY_COLOR);
        dataTable.getTableHeader().setBorder(BorderFactory.createEmptyBorder());


        dataTable.setAutoResizeMode(JTable.AUTO_RESIZE_OFF);
        for (int i = 0; i < dataTable.getColumnCount(); i++) {
            dataTable.getColumnModel().getColumn(i).setPreferredWidth(150); // 设置每列的宽度为150
        }

        return dataTable;
    }


    private JButton createStyledButton(String text) {
        JButton button = new JButton(text);
        button.setFont(new Font("Segoe UI", Font.BOLD, 12));
        button.setForeground(PRIMARY_COLOR);
        button.setBackground(CARD_BACKGROUND);


        button.setBorder(BorderFactory.createCompoundBorder(
                BorderFactory.createLineBorder(PRIMARY_COLOR, 1),
                BorderFactory.createEmptyBorder(4, 8, 4, 8)
        ));
        button.setFocusPainted(false);


        button.addMouseListener(new MouseAdapter() {
            public void mouseEntered(MouseEvent e) {
                button.setBackground(new Color(44, 62, 80, 20));
            }
            public void mouseExited(MouseEvent e) {
                button.setBackground(CARD_BACKGROUND);
            }
        });

        return button;
    }



    class ShadowBorder extends AbstractBorder {

        // Override the paintBorder method to draw the border.
        @Override
        public void paintBorder(Component c, Graphics g, int x, int y, int width, int height) {

            Graphics2D g2 = (Graphics2D) g.create();


            g2.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);


            g2.setColor(new Color(0, 0, 0, 20));

            g2.fillRoundRect(x + 2, y + 2, width - 4, height - 4, 15, 15);

            g2.dispose();
        }

        @Override
        public Insets getBorderInsets(Component c) {
            return new Insets(4, 4, 4, 4);
        }
    }

    // Main method, program entry
    public static void main(String[] args) {
        // Use SwingUtilities.invokeLater to ensure that the interface is started in an event dispatch thread
        SwingUtilities.invokeLater(() -> {
            // Get the current working directory and splice out the path to the data files
            String excelPath = System.getProperty("user.dir") + "/data.xlsx";

            // Create an instance of OptimizedDashboard, passing in the path to the file
            OptimizedDashboard dashboard = new OptimizedDashboard(excelPath);

            // Settings dashboard is visible
            dashboard.setVisible(true);
        });
    }

}

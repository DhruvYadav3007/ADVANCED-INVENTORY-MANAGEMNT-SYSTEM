package com.example;



import javax.swing.table.DefaultTableModel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.*;

import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.sql.*;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.function.BiFunction;

import javax.swing.*;

public class GUI4 implements ActionListener {
    private JFrame frame;
    private JTable table;
    private DefaultTableModel model;
    private JButton printBtn, purchaseBtn, additionalBtn, saleBtn, deleteBtn, updateBtn, FilterBtn,predictBtn,exportBtn,historyBtn;
    private JTextField usernameField;
    private JPasswordField passwordField;
    private JButton loginButton;
    private boolean loggedIn = false;

    public GUI4() {
        createLoginWindow();
    }

    private void createLoginWindow() {
        frame = new JFrame("Login");
        frame.setSize(300, 180); // Increased height for more space
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
    
        // Create a JPanel with a custom background color
        JPanel panel = new JPanel();
        panel.setBackground(new Color(52, 73, 94)); // Dark blue background
        panel.setLayout(new GridBagLayout());
        GridBagConstraints gbc = new GridBagConstraints();
        gbc.insets = new Insets(10, 10, 10, 10); // Increased insets for more spacing
    
        // Username label and field
        gbc.gridx = 0;
        gbc.gridy = 0;
        JLabel usernameLabel = new JLabel("Username:");
        usernameLabel.setForeground(Color.WHITE); // White text color
        panel.add(usernameLabel, gbc);
    
        gbc.gridx = 1;
        gbc.gridy = 0;
        usernameField = new JTextField(15);
        panel.add(usernameField, gbc);
    
        // Password label and field
        gbc.gridx = 0;
        gbc.gridy = 1;
        JLabel passwordLabel = new JLabel("Password:");
        passwordLabel.setForeground(Color.WHITE); // White text color
        panel.add(passwordLabel, gbc);
    
        gbc.gridx = 1;
        gbc.gridy = 1;
        passwordField = new JPasswordField(15);
        panel.add(passwordField, gbc);
    
        // Login button
        gbc.gridx = 0;
        gbc.gridy = 2;
        gbc.gridwidth = 2;
        gbc.anchor = GridBagConstraints.CENTER;
        loginButton = new JButton("Login");
        loginButton.setBackground(new Color(34, 167, 240)); // Light blue background
        loginButton.setForeground(Color.WHITE); // White text color
        loginButton.addActionListener(this);
        panel.add(loginButton, gbc);
    
        frame.add(panel);
        frame.setVisible(true);
    }
    
    

    private void createMainWindow() {
        frame = new JFrame("Inventory Management");
        frame.setSize(1000, 600);
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.setLayout(new BorderLayout());

        model = new DefaultTableModel();
        table = new JTable(model);
        table.setRowHeight(75);
        table.setFont(new Font("Calibri", Font.PLAIN, 20));
        JScrollPane scrollPane = new JScrollPane(table);
        frame.add(scrollPane, BorderLayout.CENTER);

        JPanel buttonPanel = new JPanel(new GridLayout(2, 4));
        buttonPanel.setPreferredSize(new Dimension(800, 100));

        printBtn = new JButton("Print");
        printBtn.addActionListener(this);
        //buttonPanel.add(printBtn);

        purchaseBtn = new JButton("Purchase");
        purchaseBtn.addActionListener(this);
      
        purchaseBtn.setForeground(Color.WHITE); 
        buttonPanel.add(purchaseBtn);
        
        additionalBtn = new JButton("Additional Purchase");
        additionalBtn.addActionListener(this);
        
        additionalBtn.setForeground(Color.WHITE); 
        buttonPanel.add(additionalBtn);
        
        saleBtn = new JButton("Sale");
        saleBtn.addActionListener(this);
       
        saleBtn.setForeground(Color.WHITE); 
        buttonPanel.add(saleBtn);
        
        deleteBtn = new JButton("Delete");
        deleteBtn.addActionListener(this);
        
        deleteBtn.setForeground(Color.WHITE); 
        buttonPanel.add(deleteBtn);
        
        updateBtn = new JButton("Update");
        updateBtn.addActionListener(this);
        
        updateBtn.setForeground(Color.WHITE); 
        buttonPanel.add(updateBtn);
        
        FilterBtn = new JButton("Filter");
        FilterBtn.addActionListener(this);
        
        FilterBtn.setForeground(Color.WHITE); 
        buttonPanel.add(FilterBtn);

        historyBtn = new JButton("History");
        historyBtn.addActionListener(this);
       
        historyBtn.setForeground(Color.WHITE); 
        buttonPanel.add(historyBtn);

        exportBtn = new JButton("Export to Excel");
        exportBtn.addActionListener(this);
       
        exportBtn.setForeground(Color.WHITE); 
        buttonPanel.add(exportBtn);


        additionalBtn.setBackground(new Color(255, 193, 7));    // Vivid Yellow (#FFC107)
        purchaseBtn.setBackground(new Color(255, 87, 34));      // Deep Orange (#FF5722)
saleBtn.setBackground(new Color(76, 175, 80));          // Green (#4CAF50)
deleteBtn.setBackground(new Color(244, 67, 54));        // Red (#F44336)
updateBtn.setBackground(new Color(0, 150, 136));        // Teal (#009688)
FilterBtn.setBackground(new Color(33, 150, 243));       // Bright Blue (#2196F3)
historyBtn.setBackground(new Color(156, 39, 176));      // Vibrant Purple (#9C27B0)
exportBtn.setBackground(new Color(0, 188, 212));        // Cyan (#00BCD4)

        frame.add(buttonPanel, BorderLayout.SOUTH);

        frame.setVisible(true);
        printBtn.doClick();
    }

    private void createUserWindow() {
        frame = new JFrame("Inventory Management");
        frame.setSize(1000, 600);
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.setLayout(new BorderLayout());

        model = new DefaultTableModel();
        table = new JTable(model);
        table.setRowHeight(75);
        table.setFont(new Font("Calibri", Font.PLAIN, 20));
        JScrollPane scrollPane = new JScrollPane(table);
        frame.add(scrollPane, BorderLayout.CENTER);

        JPanel buttonPanel = new JPanel(new GridLayout(2, 3));
        buttonPanel.setPreferredSize(new Dimension(800, 100));

        printBtn = new JButton("Print");
        printBtn.addActionListener(this);
        //buttonPanel.add(printBtn);

        purchaseBtn = new JButton("Purchase");
        purchaseBtn.addActionListener(this);
       
        purchaseBtn.setForeground(Color.WHITE); 
        buttonPanel.add(purchaseBtn);
        
        additionalBtn = new JButton("Additional Purchase");
        additionalBtn.addActionListener(this);
        
        additionalBtn.setForeground(Color.WHITE); 
        buttonPanel.add(additionalBtn);
        
        saleBtn = new JButton("Sale");
        saleBtn.addActionListener(this);
        
        saleBtn.setForeground(Color.WHITE); 
        buttonPanel.add(saleBtn);

        FilterBtn = new JButton("Filter");
        FilterBtn.addActionListener(this);
       
        FilterBtn.setForeground(Color.WHITE); 
        buttonPanel.add(FilterBtn);

        predictBtn = new JButton("Predict");
        predictBtn.addActionListener(this);
        
        predictBtn.setForeground(Color.WHITE); 
        buttonPanel.add(predictBtn);


        exportBtn = new JButton("Export to Excel");
        exportBtn.addActionListener(this);
        
        exportBtn.setForeground(Color.WHITE); 
        buttonPanel.add(exportBtn);
        
purchaseBtn.setBackground(new Color(255, 87, 34));      // Deep Orange (#FF5722)
additionalBtn.setBackground(new Color(255, 193, 7));    // Vivid Yellow (#FFC107)
saleBtn.setBackground(new Color(76, 175, 80));          // Green (#4CAF50)
FilterBtn.setBackground(new Color(33, 150, 243));       // Bright Blue (#2196F3)
predictBtn.setBackground(new Color(156, 39, 176));      // Vibrant Purple (#9C27B0)
exportBtn.setBackground(new Color(0, 188, 212));        // Cyan (#00BCD4)




   







        frame.add(buttonPanel, BorderLayout.SOUTH);
        frame.setVisible(true);
        printBtn.doClick();
    }

    public static void main(String[] args) {
        new GUI4();
    }

    @Override
    public void actionPerformed(ActionEvent e) {
      if (!loggedIn) {
    String username = usernameField.getText();
    String password = new String(passwordField.getPassword());

    if (username.equals("admin") && Encrypt.verifyAdminPassword(password)) {
        loggedIn = true;
        frame.dispose();
        createMainWindow();
    } 
    else if (username.equals("user") && Encrypt.verifyUserPassword(password)) {
        loggedIn = true;
        frame.dispose();
        createUserWindow();
    } 
    else {
        JOptionPane.showMessageDialog(frame, "Invalid username or password",
            "Login Failed", JOptionPane.ERROR_MESSAGE);
    }




        } else {

            try {
                Connection con = DriverManager.getConnection("jdbc:mysql://localhost:3306/inventory", "Dhruv", "dhruv");
                Statement stmt = con.createStatement();
    
                if (e.getSource() == printBtn) {
                    model.setColumnIdentifiers(new String[]{"Product_Id", "Product_Name", "Price", "Quantity", "Total_Purchase", "Total_Sale"});
                    model.setRowCount(0);
                    ResultSet rs = stmt.executeQuery("SELECT * FROM InventoyManagementTable");
                    while (rs.next()) {
                        Object[] row = {rs.getInt(1), rs.getString(2), rs.getInt(3), rs.getInt(4), rs.getInt(5), rs.getInt(6)};
                        model.addRow(row);
                    }
                } if (e.getSource() == purchaseBtn) {
                    if (e.getSource() == purchaseBtn) {
    JPanel panel = new JPanel(new GridLayout(0, 2));
    panel.add(new JLabel("Product ID:"));
    JTextField pidField = new JTextField();
    panel.add(pidField);

    panel.add(new JLabel("Product Name:"));
    JTextField pnameField = new JTextField();
    panel.add(pnameField);

    panel.add(new JLabel("Price:"));
    JTextField priceField = new JTextField();
    panel.add(priceField);

    panel.add(new JLabel("Quantity:"));
    JTextField qtyField = new JTextField();
    panel.add(qtyField);

    int result = JOptionPane.showConfirmDialog(null, panel, "Enter New Product Information", JOptionPane.OK_CANCEL_OPTION);

    if (result == JOptionPane.OK_OPTION) {
        int pid = Integer.parseInt(pidField.getText());
        String pname = pnameField.getText();
        int price = Integer.parseInt(priceField.getText());
        int qty = Integer.parseInt(qtyField.getText());

        try {
            // 1️⃣ Insert into inventory table
            String insertInventory = "INSERT INTO InventoyManagementTable(Product_Id, Product_Name, Price, Quantity, Total_Purchase, Total_Sale) VALUES (?, ?, ?, ?, ?, ?)";
            PreparedStatement psInventory = con.prepareStatement(insertInventory);
            psInventory.setInt(1, pid);
            psInventory.setString(2, pname);
            psInventory.setInt(3, price);
            psInventory.setInt(4, qty);
            psInventory.setInt(5, price * qty);  // Total purchase
            psInventory.setInt(6, 0);            // Total sale initially 0
            psInventory.executeUpdate();

            // 2️⃣ Insert into history table
            String historyQuery = "INSERT INTO ProductTransactions(Product_Id, Product_Name, Date_Of_Transaction, Quantity_Changed, Price, Total_Purchase, Total_Sale, Quantity_After, Cumulative_Purchase, Cumulative_Sale) " +
                                  "VALUES (?, ?, NOW(), ?, ?, ?, ?, ?, ?, ?)";
            PreparedStatement psHistory = con.prepareStatement(historyQuery);
            psHistory.setInt(1, pid);
            psHistory.setString(2, pname);
            psHistory.setInt(3, qty);               // Quantity_Changed = purchased qty
            psHistory.setInt(4, price);
            psHistory.setInt(5, price * qty);       // Total purchase for this transaction
            psHistory.setInt(6, 0);                 // Total sale for this transaction
            psHistory.setInt(7, qty);               // Quantity after transaction
            psHistory.setInt(8, price * qty);       // Cumulative purchase
            psHistory.setInt(9, 0);                 // Cumulative sale
            psHistory.executeUpdate();

            printBtn.doClick();
            JOptionPane.showMessageDialog(frame, "New product added successfully!", "Success", JOptionPane.INFORMATION_MESSAGE);

        } catch (SQLException ex) {
            JOptionPane.showMessageDialog(frame, "Error adding new product: " + ex.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
            ex.printStackTrace();
        }
    }
}

                } else if (e.getSource() == additionalBtn) {
                JPanel panel = new JPanel(new GridLayout(0, 2));
    panel.add(new JLabel("Product ID:"));
    JTextField pidField = new JTextField();
    panel.add(pidField);

    panel.add(new JLabel("Quantity:"));
    JTextField qtyField = new JTextField();
    panel.add(qtyField);

    int result = JOptionPane.showConfirmDialog(null, panel, "Enter Purchase Information", JOptionPane.OK_CANCEL_OPTION);
    if (result == JOptionPane.OK_OPTION) {
    int pid = Integer.parseInt(pidField.getText());
    int qty = Integer.parseInt(qtyField.getText());

    ResultSet quantityResult = stmt.executeQuery("SELECT Product_Name, Quantity, Price, Total_Purchase,Total_sale FROM InventoyManagementTable WHERE Product_Id=" + pid);
    if (quantityResult.next()) {
        String pname = quantityResult.getString("Product_Name");
        int currentQty = quantityResult.getInt("Quantity");
        int price = quantityResult.getInt("Price");
        int purchase = quantityResult.getInt("Total_Purchase");
        int sale =quantityResult.getInt("Total_Sale");

        int newPurchase = purchase + price * qty;
        int updatedQty = currentQty + qty;

      
        String updateQuery = "UPDATE InventoyManagementTable SET Quantity=" + updatedQty + ", Total_Purchase=" + newPurchase + " WHERE Product_Id=" + pid;
        int rowsAffected = stmt.executeUpdate(updateQuery);

     String historyQuery = "INSERT INTO ProductTransactions(Product_Id, Product_Name, Date_Of_Transaction, Quantity_Changed, Price, Total_Purchase, Total_Sale, Quantity_After, Cumulative_Purchase, Cumulative_Sale) " +
                      "VALUES (?, ?, NOW(), ?, ?, ?, 0, ?, ?, ?)";
PreparedStatement ps = con.prepareStatement(historyQuery);

ps.setInt(1, pid);
ps.setString(2, pname);
ps.setInt(3, qty);                       // positive for additional purchase
ps.setInt(4, price);
ps.setInt(5, qty*price);
ps.setInt(6, updatedQty);
ps.setDouble(7, newPurchase); // cumulative purchase
ps.setDouble(8, sale);                     // cumulative sale unchanged
ps.executeUpdate();


       
        printBtn.doClick();
        JOptionPane.showMessageDialog(frame, "Additional quantities purchased successfully for " + rowsAffected + " row(s).", "Success", JOptionPane.INFORMATION_MESSAGE);

    } else {
        JOptionPane.showMessageDialog(frame, "Product ID not found.", "Error", JOptionPane.ERROR_MESSAGE);
    }


                    }
                } else if (e.getSource() == saleBtn) {
    JPanel panel = new JPanel(new GridLayout(0, 2));
    panel.add(new JLabel("Product_ID:"));
    JTextField pidField = new JTextField();
    panel.add(pidField);

    panel.add(new JLabel("Quantity:"));
    JTextField qtyField = new JTextField();
    panel.add(qtyField);

    int result = JOptionPane.showConfirmDialog(null, panel, "Enter Sale Information", JOptionPane.OK_CANCEL_OPTION);
    if (result == JOptionPane.OK_OPTION) {
        int pid = Integer.parseInt(pidField.getText());
        int sellQty = Integer.parseInt(qtyField.getText());

        ResultSet quantityResult = stmt.executeQuery(
            "SELECT Quantity, Price, Total_Purchase,Total_Sale, Product_Name FROM InventoyManagementTable WHERE Product_Id=" + pid
        );

        if (quantityResult.next()) {
            int currentQty = quantityResult.getInt("Quantity");
            int price = quantityResult.getInt("Price");
            int currentSale = quantityResult.getInt("Total_Sale");
             int currentPur = quantityResult.getInt("Total_Purchase");

            String pname = quantityResult.getString("Product_Name");

            if (sellQty <= currentQty) {
                int updatedQty = currentQty - sellQty;
                int newSale = currentSale + sellQty * price;

                // Update Inventory Table
                String updateQuery = "UPDATE InventoyManagementTable SET Quantity=" + updatedQty + 
                                     ", Total_Sale=" + newSale + " WHERE Product_Id=" + pid;
                stmt.executeUpdate(updateQuery);

                // Insert into Transaction History Table
                String historyQuery = "INSERT INTO ProductTransactions(Product_Id, Product_Name, Date_Of_Transaction, Quantity_Changed, Price, Total_Purchase, Total_Sale, Quantity_After,Cumulative_Purchase,Cumulative_Sale) " +
                                      "VALUES (?, ?, NOW(), ?, ?, 0, ?, ?,?,?)";
                PreparedStatement ps = con.prepareStatement(historyQuery);
                ps.setInt(1, pid);
                ps.setString(2, pname);
                ps.setInt(3, -sellQty);
                ps.setInt(4, price);
                ps.setInt(5, sellQty * price);
                ps.setInt(6, updatedQty);
                ps.setInt(7, currentPur);
                ps.setInt(8,newSale);
                ps.executeUpdate();

                // Show sale success
                printBtn.doClick();
                JOptionPane.showMessageDialog(frame, sellQty + " units sold for Product ID " + pid + ".", "Success", JOptionPane.INFORMATION_MESSAGE);

                // Low stock warnings
                if (updatedQty <= 5) {
                    JOptionPane.showMessageDialog(frame, "Less than 5 quantity available for Product ID " + pid + ". Consider adding stock.", "LOW Quantity", JOptionPane.WARNING_MESSAGE);
                    additionalBtn.doClick(); // optionally trigger additional purchase
                } else if (updatedQty <= 10) {
                    JOptionPane.showMessageDialog(frame, "Less than 10 quantity available for Product ID " + pid + ".", "LOW Quantity", JOptionPane.INFORMATION_MESSAGE);
                }

            } else {
                JOptionPane.showMessageDialog(frame, "Not enough stock. Available: " + currentQty, "Error", JOptionPane.ERROR_MESSAGE);
            }
        } else {
            JOptionPane.showMessageDialog(frame, "Product ID not found.", "Error", JOptionPane.ERROR_MESSAGE);
        }
    }
}
                else if (e.getSource() == deleteBtn) {
                    int pid = Integer.parseInt(JOptionPane.showInputDialog("Enter Product ID to Delete:"));
                    String sql = "DELETE FROM InventoyManagementTable WHERE Product_Id=" + pid;
                    int rowsAffected = stmt.executeUpdate(sql);
                    if(rowsAffected>0){
                   // System.out.println(rowsAffected + " row(s) deleted successfully.");
                   printBtn.doClick();
                    JOptionPane.showMessageDialog(frame,rowsAffected + " row(s) deleted successfully.", "Success", JOptionPane.INFORMATION_MESSAGE);
                    }
                    else{
                        JOptionPane.showMessageDialog(frame, "Product ID not found.", "Error", JOptionPane.ERROR_MESSAGE);
                    }
                } else if (e.getSource() == updateBtn) {
                    JPanel panel = new JPanel(new GridLayout(0, 2));
                    panel.add(new JLabel("Product ID:"));
                    JTextField pidField = new JTextField();
                    panel.add(pidField);
                
                    panel.add(new JLabel("Column to be Updated:"));
                    JTextField fnameField = new JTextField();
                    panel.add(fnameField);
                
                    panel.add(new JLabel("Value to be Updated:"));
                    JTextField valField = new JTextField();
                    panel.add(valField);
                
                    int result = JOptionPane.showConfirmDialog(null, panel, "Enter Update Information", JOptionPane.OK_CANCEL_OPTION);
                
                    if (result == JOptionPane.OK_OPTION) {
                        int pid = Integer.parseInt(pidField.getText());
                        String columnName = fnameField.getText(); // Use a more descriptive name for the column name variable
                        String value = valField.getText(); // Use a more descriptive name for the value variable
                
                        String sql = "UPDATE inventoymanagementtable SET " + columnName + "='" + value + "' WHERE Product_Id= '" + pid+"'";
                
                        int rowsAffected = stmt.executeUpdate(sql);
                        if (rowsAffected > 0) {
                            printBtn.doClick();
                            JOptionPane.showMessageDialog(frame, rowsAffected + " row(s) updated successfully.", "Success", JOptionPane.INFORMATION_MESSAGE);
                           
                        } else {
                            JOptionPane.showMessageDialog(frame, "Product ID not found.", "Error", JOptionPane.ERROR_MESSAGE);
                        }
                    }
                }
                
                    else if (e.getSource() == FilterBtn) {
    try {
        JPanel inputPanel = new JPanel(new GridLayout(0, 2, 5, 5));
        inputPanel.add(new JLabel("Product Column (e.g., Quantity):"));
        JTextField categoryField = new JTextField();
        inputPanel.add(categoryField);

        inputPanel.add(new JLabel("Minimum Value:"));
        JTextField minQtyField = new JTextField();
        inputPanel.add(minQtyField);

        inputPanel.add(new JLabel("Maximum Value:"));
        JTextField maxQtyField = new JTextField();
        inputPanel.add(maxQtyField);

        int result = JOptionPane.showConfirmDialog(
                frame,
                inputPanel,
                "Enter Filter Criteria",
                JOptionPane.OK_CANCEL_OPTION
        );

        if (result == JOptionPane.OK_OPTION) {
            String pColumn = categoryField.getText().trim();
            int minQty = Integer.parseInt(minQtyField.getText().trim());
            int maxQty = Integer.parseInt(maxQtyField.getText().trim());

            DefaultTableModel filteredModel = new DefaultTableModel();
            filteredModel.setColumnIdentifiers(new String[]{
                    "Product_Id", "Product_Name", "Price", "Quantity",
                    "Total_Purchase", "Total_Sale"
            });

            String sql = "SELECT * FROM InventoyManagementTable WHERE " + pColumn +
                         " BETWEEN " + minQty + " AND " + maxQty;

            ResultSet rs = stmt.executeQuery(sql);
            while (rs.next()) {
                Object[] row = {
                        rs.getInt("Product_Id"),
                        rs.getString("Product_Name"),
                        rs.getInt("Price"),
                        rs.getInt("Quantity"),
                        rs.getInt("Total_Purchase"),
                        rs.getInt("Total_Sale")
                };
                filteredModel.addRow(row);
            }

            JTable filteredTable = new JTable(filteredModel);
            filteredTable.setRowHeight(30);
            filteredTable.setFont(new Font("Calibri", Font.PLAIN, 16));
            JScrollPane scrollPane = new JScrollPane(filteredTable);
            scrollPane.setPreferredSize(new Dimension(900, 400));

            // ✅ Create Export button inside Filter Result window
            JButton exportBtn = new JButton("Export to Excel");
            exportBtn.setFont(new Font("Segoe UI", Font.BOLD, 14));
            exportBtn.setPreferredSize(new Dimension(200, 35));

            exportBtn.addActionListener(ev -> {
                try {
                    JFileChooser fileChooser = new JFileChooser();
                    fileChooser.setDialogTitle("Save Excel File");
                    int resultExport = fileChooser.showSaveDialog(frame);

                    if (resultExport == JFileChooser.APPROVE_OPTION) {
                        File file = new File(fileChooser.getSelectedFile().getAbsolutePath() + ".xlsx");
                        Workbook workbook = new XSSFWorkbook();
                        Sheet sheet = workbook.createSheet("Filtered Data");

                        // Header row
                        Row headerRow = sheet.createRow(0);
                        for (int i = 0; i < filteredTable.getColumnCount(); i++) {
                            Cell cell = headerRow.createCell(i);
                            cell.setCellValue(filteredTable.getColumnName(i));
                        }

                        // Data rows
                        for (int i = 0; i < filteredTable.getRowCount(); i++) {
                            Row row = sheet.createRow(i + 1);
                            for (int j = 0; j < filteredTable.getColumnCount(); j++) {
                                Object value = filteredTable.getValueAt(i, j);
                                row.createCell(j).setCellValue(value != null ? value.toString() : "");
                            }
                        }

                        FileOutputStream fileOut = new FileOutputStream(file);
                        workbook.write(fileOut);
                        fileOut.close();
                        workbook.close();
                        JOptionPane.showMessageDialog(frame, "Data exported successfully to " + file.getAbsolutePath());
                    }
                } catch (Exception ex) {
                    JOptionPane.showMessageDialog(frame, "Error exporting data: " + ex.getMessage());
                }
            });

            JPanel resultPanel = new JPanel(new BorderLayout(10, 10));
            resultPanel.add(scrollPane, BorderLayout.CENTER);

            JPanel btnPanel = new JPanel(new FlowLayout(FlowLayout.CENTER));
            btnPanel.add(exportBtn);
            resultPanel.add(btnPanel, BorderLayout.SOUTH);

            JOptionPane.showMessageDialog(
                    frame,
                    resultPanel,
                    "Filtered Inventory Results",
                    JOptionPane.PLAIN_MESSAGE
            );
        }

    } catch (Exception ex) {
        JOptionPane.showMessageDialog(frame, "Filter error: " + ex.getMessage());
    }
}

                
                else if(e.getSource() == predictBtn ){
          
                 try {
    // Step 1: Dropdown options for available numeric columns
    String[] numericColumns = {"Price", "Quantity", "Total_Purchase", "Total_Sale"};

    // Step 2: Let user select independent (X) column
    String independentColumn = (String) JOptionPane.showInputDialog(
            frame,
            "Select the independent variable (X):",
            "Select Independent Variable",
            JOptionPane.PLAIN_MESSAGE,
            null,
            numericColumns,
            "Price"
    );

    if (independentColumn == null) return; // user cancelled

    // Step 3: Let user select dependent (Y) column
    String dependentColumn = (String) JOptionPane.showInputDialog(
            frame,
            "Select the dependent variable (Y):",
            "Select Dependent Variable",
            JOptionPane.PLAIN_MESSAGE,
            null,
            numericColumns,
            "Total_Sale"
    );

    if (dependentColumn == null) return; // user cancelled

    // Prevent same column for X and Y
    if (independentColumn.equals(dependentColumn)) {
        JOptionPane.showMessageDialog(frame, "Independent and dependent columns cannot be the same!");
        return;
    }

    // Step 4: Ask user for X input value
    String userInput = JOptionPane.showInputDialog(frame, "Enter the value for " + independentColumn + ":");
    if (userInput == null || userInput.isEmpty()) {
        JOptionPane.showMessageDialog(frame, "Please enter a valid number.");
        return;
    }

    double inputValue = Double.parseDouble(userInput);

    // Step 5: Fetch data from your Inventory database table
    String query = "SELECT " + independentColumn + ", " + dependentColumn + " FROM InventoyManagementTable";

    ArrayList<Double> independentValues = new ArrayList<>();
    ArrayList<Double> dependentValues = new ArrayList<>();

    ResultSet rs = stmt.executeQuery(query);
    while (rs.next()) {
        double independentValue = rs.getDouble(independentColumn);
        double dependentValue = rs.getDouble(dependentColumn);
        independentValues.add(independentValue);
        dependentValues.add(dependentValue);
    }

    // Step 6: Validate data
    if (independentValues.size() < 2) {
        JOptionPane.showMessageDialog(frame, "Not enough numeric data for prediction.");
        return;
    }

    // Step 7: Perform simple linear regression
    int n = independentValues.size();
    double sumX = 0, sumY = 0, sumXY = 0, sumX2 = 0;

    for (int i = 0; i < n; i++) {
        sumX += independentValues.get(i);
        sumY += dependentValues.get(i);
        sumXY += independentValues.get(i) * dependentValues.get(i);
        sumX2 += independentValues.get(i) * independentValues.get(i);
    }

    double slope = (n * sumXY - sumX * sumY) / (n * sumX2 - sumX * sumX);
    double intercept = (sumY - slope * sumX) / n;

    // Step 8: Predict Y for the given X
    double predictedValue = intercept + (slope * inputValue);

    // Step 9: Show result
    JOptionPane.showMessageDialog(frame,
            "Predicted " + dependentColumn + " for " + independentColumn + " = " +
                    inputValue + " is: " + String.format("%.2f", predictedValue));

} catch (NumberFormatException ex) {
    JOptionPane.showMessageDialog(frame, "Please enter a valid number.");
} catch (SQLException ex) {
    JOptionPane.showMessageDialog(frame, "Database query failed: " + ex.getMessage());
} catch (Exception ex) {
    JOptionPane.showMessageDialog(frame, "Prediction error: " + ex.getMessage());
}
                }

                else if(e.getSource() == exportBtn){
                     try {
            JFileChooser fileChooser = new JFileChooser();
            fileChooser.setDialogTitle("Save Excel File");
            int result = fileChooser.showSaveDialog(frame);
            if (result == JFileChooser.APPROVE_OPTION) {
                File file = new File(fileChooser.getSelectedFile().getAbsolutePath() + ".xlsx");
                Workbook workbook = new XSSFWorkbook();
                Sheet sheet = workbook.createSheet("Data Export");

                Row headerRow = sheet.createRow(0);
                for (int i = 0; i < table.getColumnCount(); i++) {
                    Cell cell = headerRow.createCell(i);
                    cell.setCellValue(table.getColumnName(i));
                }

                for (int i = 0; i < table.getRowCount(); i++) {
                    Row row = sheet.createRow(i + 1);
                    for (int j = 0; j < table.getColumnCount(); j++) {
                        Object value = table.getValueAt(i, j);
                        row.createCell(j).setCellValue(value != null ? value.toString() : "");
                    }
                }

                FileOutputStream fileOut = new FileOutputStream(file);
                workbook.write(fileOut);
                fileOut.close();
                workbook.close();
                JOptionPane.showMessageDialog(frame, "Data exported successfully to " + file.getAbsolutePath());
            }
        } catch (Exception ex) {
            JOptionPane.showMessageDialog(frame, "Error exporting data: " + ex.getMessage());
        }
                }
            

            else if(e.getSource() == historyBtn){
                // ================= Transaction History Frame ===================
JFrame historyFrame = new JFrame("Transaction History");
historyFrame.setDefaultCloseOperation(JFrame.DISPOSE_ON_CLOSE);
historyFrame.setExtendedState(JFrame.MAXIMIZED_BOTH);

// Table Model
DefaultTableModel model = new DefaultTableModel();
String[] columns = {"Transaction_Id","Product_Id","Product_Name","Date_of_Transaction",
                    "Quantity_Changed","Price","Total_Purchase","Total_Sale",
                    "Quantity_After","Cumulative_Purchase","Cumulative_Sale"};
model.setColumnIdentifiers(columns);

// Load all transaction data
String sql = "SELECT * FROM ProductTransactions";
ResultSet rs = stmt.executeQuery(sql);
while(rs.next()) {
    Object[] row = {
        rs.getInt("Transaction_Id"),
        rs.getInt("Product_Id"),
        rs.getString("Product_Name"),
        rs.getTimestamp("Date_Of_Transaction"),
        rs.getInt("Quantity_Changed"),
        rs.getDouble("Price"),
        rs.getDouble("Total_Purchase"),
        rs.getDouble("Total_Sale"),
        rs.getInt("Quantity_After"),
        rs.getDouble("Cumulative_Purchase"),
        rs.getDouble("Cumulative_Sale")
    };
    model.addRow(row);
}

// JTable
JTable table = new JTable(model);
//table.setAutoResizeMode(JTable.AUTO_RESIZE_OFF);
table.setFont(new Font("Calibri", Font.PLAIN, 16));
table.setRowHeight(30);
JScrollPane scrollPane = new JScrollPane(table,
        JScrollPane.VERTICAL_SCROLLBAR_AS_NEEDED,
        JScrollPane.HORIZONTAL_SCROLLBAR_AS_NEEDED);

// Buttons panel

 JPanel buttonPanel = new JPanel(new GridLayout(1, 2));
        buttonPanel.setPreferredSize(new Dimension(200,40));
JButton filterBtn = new JButton("Filter");

JButton exportBtn = new JButton("Export");
filterBtn.setBackground(new Color(76, 175, 80));          // Green (#4CAF50)
exportBtn.setBackground(new Color(244, 67, 54));   
buttonPanel.add(filterBtn);

buttonPanel.add(exportBtn);

historyFrame.getContentPane().setLayout(new BorderLayout());
historyFrame.getContentPane().add(scrollPane, BorderLayout.CENTER);
historyFrame.getContentPane().add(buttonPanel, BorderLayout.SOUTH);
historyFrame.setVisible(true);


exportBtn.addActionListener(ev -> {
                try {
                    JFileChooser fileChooser = new JFileChooser();
                    fileChooser.setDialogTitle("Save Excel File");
                    int res = fileChooser.showSaveDialog(historyFrame);
                    if(res == JFileChooser.APPROVE_OPTION){
                        File file = new File(fileChooser.getSelectedFile().getAbsolutePath()+".xlsx");
                        Workbook workbook = new XSSFWorkbook();
                        Sheet sheet = workbook.createSheet("Filtered Data");

                        Row headerRow = sheet.createRow(0);
                        for(int i=0;i<table.getColumnCount();i++){
                            Cell cell = headerRow.createCell(i);
                            cell.setCellValue(table.getColumnName(i));
                        }

                        for(int i=0;i<table.getRowCount();i++){
                            Row row = sheet.createRow(i+1);
                            for(int j=0;j<table.getColumnCount();j++){
                                Object value = table.getValueAt(i,j);
                                row.createCell(j).setCellValue(value!=null?value.toString():"");
                            }
                        }
                        FileOutputStream out = new FileOutputStream(file);
                        workbook.write(out);
                        out.close();
                        workbook.close();
                        JOptionPane.showMessageDialog(historyFrame,"Data exported to "+file.getAbsolutePath());
                    }
                } catch(Exception exe) {
                    JOptionPane.showMessageDialog(historyFrame,"Export error: "+exe.getMessage());
                }
            });
// ================= FILTER BUTTON ===================
       filterBtn.addActionListener(ex -> {
    try {
        JPanel panel = new JPanel(new GridLayout(0,2));
        panel.add(new JLabel("Product Name (leave blank for all):"));
        JTextField productField = new JTextField();
        panel.add(productField);

        panel.add(new JLabel("Column to filter:"));
        JTextField columnField = new JTextField();
        panel.add(columnField);

        panel.add(new JLabel("Min Value:"));
        JTextField minField = new JTextField();
        panel.add(minField);

        panel.add(new JLabel("Max Value:"));
        JTextField maxField = new JTextField();
        panel.add(maxField);

        panel.add(new JLabel("From Date (yyyy-MM-dd):"));
        JTextField fromDateField = new JTextField();
        panel.add(fromDateField);

        panel.add(new JLabel("To Date (yyyy-MM-dd):"));
        JTextField toDateField = new JTextField();
        panel.add(toDateField);

        int result = JOptionPane.showConfirmDialog(historyFrame, panel, "Filter Criteria", JOptionPane.OK_CANCEL_OPTION);
        if(result == JOptionPane.OK_OPTION) {
            String product = productField.getText().trim();
            String column = columnField.getText().trim();
            String min = minField.getText().trim();
            String max = maxField.getText().trim();
            String fromDate = fromDateField.getText().trim();
            String toDate = toDateField.getText().trim();

            StringBuilder query = new StringBuilder("SELECT * FROM ProductTransactions WHERE 1=1");

            if(!product.isEmpty()) 
                query.append(" AND Product_Name = '").append(product).append("'");
            
            if(!column.isEmpty() && !min.isEmpty() && !max.isEmpty()) 
                query.append(" AND ").append(column).append(" BETWEEN ").append(min).append(" AND ").append(max);

            // ✅ Add optional date filter
            if(!fromDate.isEmpty() && !toDate.isEmpty())
                query.append(" AND Date_Of_Transaction BETWEEN '").append(fromDate).append("' AND '").append(toDate).append("'");

            ResultSet rsFiltered = stmt.executeQuery(query.toString());

            DefaultTableModel filteredModel = new DefaultTableModel();
            filteredModel.setColumnIdentifiers(columns);

            while(rsFiltered.next()) {
                Object[] row = {
                    rsFiltered.getInt("Transaction_Id"),
                    rsFiltered.getInt("Product_Id"),
                    rsFiltered.getString("Product_Name"),
                    rsFiltered.getTimestamp("Date_Of_Transaction"),
                    rsFiltered.getInt("Quantity_Changed"),
                    rsFiltered.getDouble("Price"),
                    rsFiltered.getDouble("Total_Purchase"),
                    rsFiltered.getDouble("Total_Sale"),
                    rsFiltered.getInt("Quantity_After"),
                    rsFiltered.getDouble("Cumulative_Purchase"),
                    rsFiltered.getDouble("Cumulative_Sale")
                };
                filteredModel.addRow(row);
            }

            // Show filtered table in new JFrame
            JFrame filteredFrame = new JFrame("Filtered Transactions");
            filteredFrame.setDefaultCloseOperation(JFrame.DISPOSE_ON_CLOSE);
            filteredFrame.setExtendedState(JFrame.MAXIMIZED_BOTH);

            JTable filteredTable = new JTable(filteredModel);
            filteredTable.setFont(new Font("Calibri", Font.PLAIN, 16));
            filteredTable.setRowHeight(30);
            JScrollPane sp = new JScrollPane(filteredTable,
                    JScrollPane.VERTICAL_SCROLLBAR_AS_NEEDED,
                    JScrollPane.HORIZONTAL_SCROLLBAR_AS_NEEDED);


            // Buttons on filtered frame
            JPanel filteredBtnPanel = new JPanel(new GridLayout(1, 2));

            JButton predictBtnFiltered = new JButton("Predict");
           
        predictBtnFiltered.setPreferredSize(new Dimension(200, 40));
            JButton exportBtnFiltered = new JButton("Export");
            predictBtnFiltered.setBackground(new Color(76, 175, 80));          // Green (#4CAF50)
exportBtnFiltered.setBackground(new Color(244, 67, 54));   
            filteredBtnPanel.add(predictBtnFiltered);
            filteredBtnPanel.add(exportBtnFiltered);

            filteredFrame.getContentPane().setLayout(new BorderLayout());
            filteredFrame.getContentPane().add(sp, BorderLayout.CENTER);
            filteredFrame.getContentPane().add(filteredBtnPanel, BorderLayout.SOUTH);
            filteredFrame.setVisible(true);

            // ================= PREDICT BUTTON ===================
    predictBtnFiltered.addActionListener(ev -> {
    try {
        String[] predictionOptions = {
                "1. Date when stock will run out",
                "2. Cumulative Revenue Forecast",
                "3. Sales Trend Analysis",
                "4. Future Stock after next predicted sale"
        };

        String selectedOption = (String) JOptionPane.showInputDialog(
                filteredFrame,
                "Select Prediction Type:",
                "Prediction Type",
                JOptionPane.PLAIN_MESSAGE,
                null,
                predictionOptions,
                predictionOptions[0]
        );
        if (selectedOption == null) return;

        ArrayList<Double> xDays = new ArrayList<>();
        ArrayList<Double> quantityAfter = new ArrayList<>();
        ArrayList<Double> cumulativeSale = new ArrayList<>();
        ArrayList<Double> quantityChangedSales = new ArrayList<>();

        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
        Date firstDate = null;

        String productName = filteredTable.getValueAt(0, 2).toString(); // Product_Name

        for (int i = 0; i < filteredTable.getRowCount(); i++) {
            Timestamp ts = (Timestamp) filteredTable.getValueAt(i, 3); // Date_of_Transaction
            if (firstDate == null) firstDate = new Date(ts.getTime());
            long diffDays = (ts.getTime() - firstDate.getTime()) / (1000 * 60 * 60 * 24);
            xDays.add((double) diffDays);

            int qtyAfter = Integer.parseInt(filteredTable.getValueAt(i, 8).toString()); // Quantity_After
            quantityAfter.add((double) qtyAfter);

            double cumSale = Double.parseDouble(filteredTable.getValueAt(i, 10).toString()); // Cumulative_Sale
            cumulativeSale.add(cumSale);

            int qtyChanged = Integer.parseInt(filteredTable.getValueAt(i, 4).toString()); // Quantity_Changed
            if (qtyChanged < 0) quantityChangedSales.add((double) (-qtyChanged)); // sales only
        }

        if (xDays.size() < 2) {
            JOptionPane.showMessageDialog(filteredFrame, "Not enough data for prediction.");
            return;
        }

        // Linear regression function
        BiFunction<ArrayList<Double>, ArrayList<Double>, double[]> linearRegression = (xs, ys) -> {
            int n = xs.size();
            double sumX = 0, sumY = 0, sumXY = 0, sumX2 = 0;
            for (int i = 0; i < n; i++) {
                sumX += xs.get(i);
                sumY += ys.get(i);
                sumXY += xs.get(i) * ys.get(i);
                sumX2 += xs.get(i) * xs.get(i);
            }
            double slope = (n * sumXY - sumX * sumY) / (n * sumX2 - sumX * sumX);
            double intercept = (sumY - slope * sumX) / n;
            return new double[]{slope, intercept};
        };

        String message = "";

        switch (selectedOption.charAt(0)) {
     case '1': // Predict stock run-out
    try {
        // 1️⃣ Collect sales-only quantities and dates
        ArrayList<Double> salesDays = new ArrayList<>();
        ArrayList<Double> salesQty = new ArrayList<>();
        Date lastDate = null;
        int currentStock = 0;

        for (int i = 0; i < filteredTable.getRowCount(); i++) {
            Timestamp ts = (Timestamp) filteredTable.getValueAt(i, 3); // Date_of_Transaction
            int qtyChanged = Integer.parseInt(filteredTable.getValueAt(i, 4).toString());
            int qtyAfter = Integer.parseInt(filteredTable.getValueAt(i, 8).toString());
            
            if (lastDate == null || ts.after(lastDate)) {
                lastDate = new Date(ts.getTime());
                currentStock = qtyAfter; // current stock = last Quantity_After
            }

            if (qtyChanged < 0) { // only sales
                salesDays.add((double) i); // simple day index for regression
                salesQty.add((double) (-qtyChanged));
            }
        }

        if (salesQty.size() < 1) {
            message = "No sales detected for '" + productName + "' — stock will not run out soon.";
            break;
        }

        // 2️⃣ Compute average daily sales
        double totalSales = salesQty.stream().mapToDouble(Double::doubleValue).sum();
        double daysSpan = salesDays.get(salesDays.size() - 1) - salesDays.get(0) + 1;
        double avgDailySale = totalSales / daysSpan;

        if (avgDailySale <= 0) {
            message = "Average daily sales is zero — stock will not run out soon.";
            break;
        }

        // 3️⃣ Ask user for target stock
        String input = JOptionPane.showInputDialog(
                filteredFrame,
                "Enter stock level (units) to predict when it will reach:",
                "0"
        );
        if (input == null || input.trim().isEmpty()) {
            message = "Prediction cancelled.";
            break;
        }
        int targetStock = Integer.parseInt(input.trim());

        // 4️⃣ Calculate days to reach target stock
        double daysToTarget = (currentStock - targetStock) / avgDailySale;
        if (daysToTarget < 0) {
            message = "Stock is already below the target level.";
            break;
        }

        // 5️⃣ Compute predicted date from last transaction date
        Calendar cal = Calendar.getInstance();
        cal.setTime(lastDate);
        cal.add(Calendar.DATE, (int) Math.ceil(daysToTarget));

        message = "Predicted date when stock of '" + productName +
                  "' will reach " + targetStock + " units: " + new SimpleDateFormat("yyyy-MM-dd").format(cal.getTime());

    } catch (Exception exe) {
        message = "Error predicting stock: " + exe.getMessage();
    }
    break;


           case '2': // Cumulative Revenue Forecast
    double[] regRevenue = linearRegression.apply(xDays, cumulativeSale);
    double slopeRev = regRevenue[0];
    double interceptRev = regRevenue[1];


    String inputDays = JOptionPane.showInputDialog(
            filteredFrame,
            "Enter number of days to forecast cumulative revenue:",
            "7"
    );
    if (inputDays == null || inputDays.trim().isEmpty()) return;
    int d2 = Integer.parseInt(inputDays.trim()); // days ahead

    double forecastRevenue = interceptRev + slopeRev * (xDays.get(xDays.size() - 1) + d2);

    message = " Predicted cumulative sales of '" + productName +
            "' after " + d2 + " days: " + String.format("%.2f", forecastRevenue);
    break;


          case '3': // Numerical Sales Prediction
    if (xDays.size() < 2) {
        message = "Not enough data for sales prediction.";
        break;
    }

    // Build daily sales list (convert negative sales to positive numbers)
    ArrayList<Double> dailySales = new ArrayList<>();
    for (int i = 0; i < xDays.size(); i++) {
        int qtyChanged = Integer.parseInt(filteredTable.getValueAt(i, 4).toString());
        double sale = qtyChanged < 0 ? -qtyChanged : 0; // ignore restocks
        dailySales.add(sale);
    }

    double[] regSales = linearRegression.apply(xDays, dailySales);
    double slopeSales = regSales[0];
    double interceptSales = regSales[1];

    // Ask how many future days to forecast sales for
    String inputDays3 = JOptionPane.showInputDialog(
            filteredFrame,
            "Enter number of future days to forecast sales:",
            "7"
    );
    if (inputDays3 == null || inputDays3.trim().isEmpty()) return;
    int d3 = Integer.parseInt(inputDays3.trim());

    double futureX3 = xDays.get(xDays.size() - 1) + d3;
    double predictedSales = interceptSales + slopeSales * futureX3;
    double avgSale = dailySales.stream().mapToDouble(Double::doubleValue).average().orElse(0);

    message = " Predicted daily sales of '" + productName + "' after " + d3 +
            " days: " + String.format("%.2f", predictedSales) +
            "\nAverage historical sales per day: " + String.format("%.2f", avgSale) +
            "\nTrend slope: " + String.format("%.4f", slopeSales);
    break;

case '4': // Future Stock Forecast
    try {
        if (xDays.size() < 2) {
            message = "Not enough data to predict future stock.";
            break;
        }

        // Ask number of days ahead
        String inputDays4 = JOptionPane.showInputDialog(
                filteredFrame,
                "Enter number of future days to predict stock for:",
                "1"
        );
        if (inputDays4 == null || inputDays4.trim().isEmpty()) {
            message = "Prediction cancelled.";
            break;
        }
        int d4 = Integer.parseInt(inputDays4.trim());

        // Compute current stock and average daily sales
        int currentStock = Integer.parseInt(filteredTable.getValueAt(filteredTable.getRowCount()-1, 8).toString());
        ArrayList<Double> salesDays = new ArrayList<>();
        ArrayList<Double> salesQty = new ArrayList<>();
        Date lastDate = null;

        for (int i = 0; i < filteredTable.getRowCount(); i++) {
            Timestamp ts = (Timestamp) filteredTable.getValueAt(i, 3);
            int qtyChanged = Integer.parseInt(filteredTable.getValueAt(i, 4).toString());

            if (lastDate == null || ts.after(lastDate)) {
                lastDate = new Date(ts.getTime());
            }

            if (qtyChanged < 0) { // only sales
                salesDays.add((double)i);
                salesQty.add((double)(-qtyChanged));
            }
        }

        if (salesQty.size() < 1) {
            message = "No sales detected — cannot forecast future stock.";
            break;
        }

        // average daily sales
        double totalSales = salesQty.stream().mapToDouble(Double::doubleValue).sum();
        double daysSpan = salesDays.get(salesDays.size()-1) - salesDays.get(0) + 1;
        double avgDailySale = totalSales / daysSpan;

        // predict future stock
        double predictedStock = currentStock - avgDailySale * d4;
        if(predictedStock < 0) predictedStock = 0;

        // compute future date from last transaction
        Calendar cal4 = Calendar.getInstance();
        cal4.setTime(lastDate);
        cal4.add(Calendar.DATE, d4);

        message = "Predicted stock of '" + productName + "' after " + d4 +
                " days (on " + sdf.format(cal4.getTime()) + "): " +
                String.format("%.2f", predictedStock);

    } catch (Exception exe) {
        message = "Error predicting future stock: " + exe.getMessage();
    }
    break;


        }

        JOptionPane.showMessageDialog(filteredFrame, message, "Prediction Result", JOptionPane.INFORMATION_MESSAGE);

    } catch (Exception exe) {
        JOptionPane.showMessageDialog(filteredFrame, "Prediction error: " + exe.getMessage());
    }
});

    // ================= EXPORT BUTTON ===================
    exportBtnFiltered.addActionListener(ev -> {
                try {
                    JFileChooser fileChooser = new JFileChooser();
                    fileChooser.setDialogTitle("Save Excel File");
                    int res = fileChooser.showSaveDialog(filteredFrame);
                    if(res == JFileChooser.APPROVE_OPTION){
                        File file = new File(fileChooser.getSelectedFile().getAbsolutePath()+".xlsx");
                        Workbook workbook = new XSSFWorkbook();
                        Sheet sheet = workbook.createSheet("Filtered Data");

                        Row headerRow = sheet.createRow(0);
                        for(int i=0;i<filteredTable.getColumnCount();i++){
                            Cell cell = headerRow.createCell(i);
                            cell.setCellValue(filteredTable.getColumnName(i));
                        }

                        for(int i=0;i<filteredTable.getRowCount();i++){
                            Row row = sheet.createRow(i+1);
                            for(int j=0;j<filteredTable.getColumnCount();j++){
                                Object value = filteredTable.getValueAt(i,j);
                                row.createCell(j).setCellValue(value!=null?value.toString():"");
                            }
                        }
                        FileOutputStream out = new FileOutputStream(file);
                        workbook.write(out);
                        out.close();
                        workbook.close();
                        JOptionPane.showMessageDialog(filteredFrame,"Data exported to "+file.getAbsolutePath());
                    }
                } catch(Exception exe) {
                    JOptionPane.showMessageDialog(filteredFrame,"Export error: "+exe.getMessage());
                }
            });

        }
    } catch(Exception exe) {
        JOptionPane.showMessageDialog(historyFrame,"Filter error: "+exe.getMessage());
    }
});

            }
            
                    
                    
                    
            
            } catch (Exception ex) {
                //System.out.println(ex);
                JOptionPane.showMessageDialog(frame,ex, "Failure", JOptionPane.INFORMATION_MESSAGE);
            }
        }
    }
    
}

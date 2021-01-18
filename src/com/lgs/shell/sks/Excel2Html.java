
package com.lgs.shell.sks;

import com.lgs.workers.ConversionWorker;
import java.awt.BorderLayout;
import java.awt.Component;
import java.awt.Dimension;
import java.awt.FlowLayout;
import java.awt.Toolkit;
import java.io.File;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.JButton;
import javax.swing.JDialog;
import javax.swing.JFileChooser;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JProgressBar;
import javax.swing.JScrollPane;
import javax.swing.JTextArea;
import javax.swing.UIManager;
import javax.swing.UnsupportedLookAndFeelException;
import javax.swing.border.EmptyBorder;
import javax.swing.filechooser.FileNameExtensionFilter;

/**
 *
 * @author ShivanshuJS
 */
public class Excel2Html extends JDialog{

    // Swing UI components which are required to access throughout the class.
    private final JFileChooser fileSelector = new JFileChooser();
    private final JPanel mainPanel = new JPanel();
    private final JButton fileInvokeButton = new JButton("Select File");
    private final JButton convertButton = new JButton("Convert");
    private final JTextArea infoTextArea = new JTextArea();
    private final JProgressBar progressBar = new JProgressBar(0, 100);
    
    // Constructor to initialize Swing UI Components.
    public Excel2Html(){
        this.initComponents();
    }
    
    // Initialize the Swing UI Components and create the dialog.
    private void initComponents(){
        this.setIconImage(Toolkit.getDefaultToolkit().getImage(Excel2Html.class.getResource("/com/lgs/resources/shell.png")));
        this.setTitle("Excel To HTML Converter");
        this.setBounds(100, 100, 380, 200);
        this.getContentPane().setLayout(new BorderLayout());
        this.setResizable(false);
        this.mainPanel.setBorder(new EmptyBorder(5, 5, 5, 5));
        this.mainPanel.setLayout(new BorderLayout());
        JPanel filePanel = new JPanel(new FlowLayout(FlowLayout.LEFT));
        JLabel fileLabel = new JLabel("Select a Microsoft Excel Document: ");
        filePanel.add((Component)fileLabel);
        this.fileInvokeButton.addActionListener((e) -> {
            this.progressBar.setValue(0);
            if(this.showFileSelectorDialog() == JFileChooser.APPROVE_OPTION){
                this.infoTextArea.setText("Selected File: " + this.fileSelector.getSelectedFile().getAbsolutePath());
                this.convertButton.setEnabled(true);
            } else {
                this.infoTextArea.setText("");
                JOptionPane.showMessageDialog(null, "You have not selected any file.", "Alert", JOptionPane.ERROR_MESSAGE);
            }
        });
        filePanel.add((Component)this.fileInvokeButton);
        this.convertButton.setEnabled(false);
        this.convertButton.addActionListener((e) -> {
            this.fileInvokeButton.setEnabled(false);
            this.convertButton.setEnabled(false);
            ConversionWorker conversionWorker = new ConversionWorker(this.fileSelector.getSelectedFile().getAbsolutePath(), this.infoTextArea, this.fileInvokeButton);
            conversionWorker.addPropertyChangeListener((propertyChangeEvent) -> {
                if("progress".equals(propertyChangeEvent.getPropertyName())){
                    this.progressBar.setValue((Integer)propertyChangeEvent.getNewValue());
                }
            });
            conversionWorker.execute();
        });
        filePanel.add((Component)this.convertButton);
        this.mainPanel.add((Component)filePanel, BorderLayout.PAGE_START);
        JPanel selectedFilePanel = new JPanel(new BorderLayout());
        this.infoTextArea.setEditable(false);
        this.infoTextArea.setLineWrap(true);
        this.infoTextArea.setWrapStyleWord(true);
        this.infoTextArea.setFont(this.infoTextArea.getFont().deriveFont(11f));
        JScrollPane infoScroller = new JScrollPane(this.infoTextArea);
        infoScroller.setPreferredSize(new Dimension(340, 25));
        selectedFilePanel.add((Component)infoScroller, "Center");
        this.mainPanel.add((Component)selectedFilePanel, BorderLayout.CENTER);
        JPanel progressPanel = new JPanel(new FlowLayout(FlowLayout.LEFT));
        this.progressBar.setPreferredSize(new Dimension(345, 25));
        this.progressBar.setStringPainted(true);
        progressPanel.add((Component)this.progressBar);
        this.mainPanel.add((Component)progressPanel, BorderLayout.PAGE_END);
        this.getContentPane().add((Component)this.mainPanel, "Center");
    }
    
    // Configure the JFileChooser and show it to the user to select the file.
    private int showFileSelectorDialog(){
        this.fileSelector.setCurrentDirectory(new File(System.getProperty("user.home")));
        this.fileSelector.setFileSelectionMode(JFileChooser.FILES_ONLY);
        this.fileSelector.setAcceptAllFileFilterUsed(false);
        this.fileSelector.addChoosableFileFilter(new FileNameExtensionFilter("Microsoft Excel Documents", "xls"));
        int isFileSelected = this.fileSelector.showOpenDialog(null);
        return isFileSelected;
    }
    
    // The main method. Program execution starts from here.
    public static void main(String[] args) {
        try {
            UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
            Excel2Html wordToHtmlDialog = new Excel2Html();
            wordToHtmlDialog.setDefaultCloseOperation(2);
            wordToHtmlDialog.setVisible(true);
        } catch (ClassNotFoundException | InstantiationException | IllegalAccessException | UnsupportedLookAndFeelException ex) {
            Logger.getLogger(Excel2Html.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
}

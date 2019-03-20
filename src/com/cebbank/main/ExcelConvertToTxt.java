package com.cebbank.main;

import java.awt.Container;
import java.awt.Point;
import java.awt.Toolkit;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.IOException;

import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JTextArea;
import javax.swing.JTextField;
import javax.swing.filechooser.FileNameExtensionFilter;

import com.cebbank.util.ExcelToTxt;

public class ExcelConvertToTxt implements ActionListener {

	JFrame frame = new JFrame("Excel格式转换成txt格式");

	Container con = new Container();
	JLabel label1 = new JLabel("选择文件路径");
	JLabel label4 = new JLabel();
	JLabel label5 = new JLabel();

	JTextField text1 = new JTextField("");
	JTextArea text_result = new JTextArea();

	JButton button1 = new JButton("选择");
	JButton button5 = new JButton("转换");
	JFileChooser jfc = new JFileChooser();

	ExcelConvertToTxt() throws Exception {
		text_result.setVisible(false);
		// jfc.setCurrentDirectory(new File("d:\\"));
		jfc.setFileSelectionMode(JFileChooser.FILES_ONLY);
		FileNameExtensionFilter filter = new FileNameExtensionFilter("*.xls", "xls", "*.xlsx", "xlsx");
		jfc.setFileFilter(filter);
		jfc.setCurrentDirectory(new File("."));

		double lx = Toolkit.getDefaultToolkit().getScreenSize().getWidth();
		double ly = Toolkit.getDefaultToolkit().getScreenSize().getHeight();
		frame.setLocation(new Point((int) (lx / 2) - 230, (int) (ly / 2) - 120));// 设定窗口出现位置
		frame.setSize(600, 240);// 设定窗口大小
		// 下面设定标签等的出现位置和宽
		label1.setBounds(10, 30, 100, 20);
		text1.setBounds(110, 30, 250, 20);
		button1.setBounds(400, 30, 100, 20);
		text1.setEditable(false);

		label4.setBounds(10, 70, 1000, 20);
		label5.setBounds(100, 90, 1000, 20);

		button5.setBounds(160, 160, 100, 20);

		text_result.setBounds(10, 100, 970, 300);
		text_result.setAutoscrolls(true);

		button1.addActionListener(this);
		button5.addActionListener(this);

		con.add(label1);
		con.add(text1);
		con.add(button1);

		con.add(label4);
		con.add(label5);

		con.add(button5);
		con.add(jfc);

		con.add(text_result);

		frame.add(con);
		frame.setVisible(true);
		frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
	}

	public void actionPerformed(ActionEvent e) {
		if (e.getSource().equals(button1)) {
			label4.setText("");
			label5.setText("");
			jfc.setFileSelectionMode(JFileChooser.FILES_ONLY);
			int state = jfc.showOpenDialog(null);
			if (state == 1) {
				return;
			} else {
				File f = jfc.getSelectedFile();
				text1.setText(f.getAbsolutePath());
			}
		}

		if (e.getSource().equals(button5)) { // 导入
			if (text1.getText().trim().equals("")) {
				label4.setText("请选择文件路径！");
				label5.setText("");
			} else if(!checkFileName(text1.getText())){
				label4.setText("请上传Excel格式文件！");
				label5.setText("");
			}else{
				String txtPath = "";
				try {
					txtPath = convertFile(text1.getText());
				} catch (IOException e1) {
					label4.setText(e1.getMessage());
					e1.printStackTrace();
				}

				label4.setText("文件转换完成，新生成的txt文件路径如下：");
				label5.setText(txtPath);
			}
		}
	}
	private boolean checkFileName(String filePath) {
		String extString = filePath.substring(filePath.lastIndexOf(".")).toLowerCase();
		if(".xls".equals(extString)||(".xlsx".equals(extString))) {
			return true;
		}
		return false;
		
	}
	
	
	/**
	 * 将目标文件Excel文件转换成txt文件。
	 * @param filePath 返回生成的txt文件格式路径
	 * @return
	 * @throws IOException
	 */
	private String convertFile(String filePath) throws IOException {
		String txtPath = ExcelToTxt.convertExcel(filePath);
		return txtPath;
	}

	public static void main(String[] args) throws Exception {
		new ExcelConvertToTxt();
	}

}
package com.nari.sung;

import java.awt.Color;
import java.awt.Point;
import java.awt.Toolkit;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.util.HashMap;
import java.util.Map;

import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JTextField;

public class Transformer extends JFrame implements ActionListener {

	/**
	 * 
	 */
	private static final long serialVersionUID = 1L;

	public static void main(String[] args) {
		Transformer transformer = new Transformer();
		transformer.init();
	}

	private JLabel selectJLabel = new JLabel("选择文件:");
	private JTextField jTextField = new JTextField();
	private JLabel selectJLabel2 = new JLabel("Sheet,TitleRow:");
	private JTextField jTextField2 = new JTextField();
	private JLabel selectJLabel3 = new JLabel("承包方编码:");
	private JTextField jTextField3 = new JTextField();
	private JLabel selectJLabel4 = new JLabel("承包方地址:");
	private JTextField jTextField4 = new JTextField();
	private JLabel selectJLabel5 = new JLabel("邮政编码:");
	private JTextField jTextField5 = new JTextField();
	private JLabel selectJLabel6 = new JLabel("调查员:");
	private JTextField jTextField6 = new JTextField();
	private JLabel selectJLabel7 = new JLabel("调查日期:");
	private JTextField jTextField7 = new JTextField();
	private JButton selectJButton = new JButton("...");
	private JButton transJButton = new JButton("转换");
	private JLabel resultJLabel = new JLabel();
	private JFileChooser jFileChooser = new JFileChooser();// 文件选择器

	/**
	 * 初始化界面的方法
	 */
	private void init() {
		this.setTitle("Excel文件转换");
		// 下面两行是取得屏幕的高度和宽度
		double lx = Toolkit.getDefaultToolkit().getScreenSize().getWidth();
		double ly = Toolkit.getDefaultToolkit().getScreenSize().getHeight();
		this.setLocation(new Point((int) (lx / 2) - 150, (int) (ly / 2) - 150));// 设定窗口出现位置
		this.setSize(500, 300);
		this.setResizable(false);
		this.setLayout(null);
		this.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);

		jFileChooser.setCurrentDirectory(new File("E:\\"));// 文件选择器的初始目录定为d盘

		// 下面设定标签等的出现位置和高宽
		selectJLabel.setBounds(10, 30, 70, 20);
		jTextField.setBounds(80, 30, 280, 20);
		jTextField.setEditable(false);
		selectJButton.setBounds(360, 30, 50, 20);
		selectJButton.addActionListener(this);// 添加事件处理
		transJButton.setBounds(410, 30, 80, 20);
		transJButton.addActionListener(this);// 添加事件处理
		selectJLabel2.setBounds(10, 60, 70, 20);
		jTextField2.setBounds(80, 60, 280, 20);
		selectJLabel3.setBounds(10, 90, 70, 20);
		jTextField3.setBounds(80, 90, 280, 20);
		selectJLabel4.setBounds(10, 120, 70, 20);
		jTextField4.setBounds(80, 120, 280, 20);
		selectJLabel5.setBounds(10, 150, 70, 20);
		jTextField5.setBounds(80, 150, 280, 20);
		selectJLabel6.setBounds(10, 180, 70, 20);
		jTextField6.setBounds(80, 180, 280, 20);
		selectJLabel7.setBounds(10, 210, 70, 20);
		jTextField7.setBounds(80, 210, 280, 20);
		resultJLabel.setBounds(10, 250, 400, 20);

		this.add(selectJLabel);
		this.add(jTextField);
		this.add(selectJButton);
		this.add(transJButton);
		this.add(selectJLabel2);
		this.add(jTextField2);
		this.add(selectJLabel3);
		this.add(jTextField3);
		this.add(selectJLabel4);
		this.add(jTextField4);
		this.add(selectJLabel5);
		this.add(jTextField5);
		this.add(selectJLabel6);
		this.add(jTextField6);
		this.add(selectJLabel7);
		this.add(jTextField7);
		this.add(resultJLabel);
		this.add(jFileChooser);

		this.setVisible(true);// 窗口可见
	}

	public void actionPerformed(ActionEvent event) {// 事件处理
		if (event.getSource().equals(selectJButton)) {
			jFileChooser.setFileSelectionMode(0);// 设定只能选择到文件
			int state = jFileChooser.showOpenDialog(null);// 此句是打开文件选择器界面的触发语句
			if (state == 1) {
				return;// 撤销则返回
			} else {
				File f = jFileChooser.getSelectedFile();// f为选择到的文件
				jTextField.setText(f.getAbsolutePath());

				if (!StringUtil.isEmptyString(resultJLabel.getText())) {
					resultJLabel.setText("");
				}
			}
		}

		if (event.getSource().equals(transJButton)) {
			String fName = jTextField.getText();
			if ("".equals(fName)) {
				resultJLabel.setForeground(Color.red);
				resultJLabel.setText("请选择需要转换的文件！");
				return;
			}

			if (-1 == fName.indexOf(".xls") && -1 == fName.indexOf(".xlsx")) {
				resultJLabel.setForeground(Color.red);
				resultJLabel.setText("仅支持xls或xlsx格式文件！");
				return;
			}
			try {
				Map<String, String> map = new HashMap<String, String>();
				map.put("SR", jTextField2.getText());
				map.put("CBFBM", jTextField3.getText());
				map.put("CBFDZ", jTextField4.getText());
				map.put("YZBM", jTextField5.getText());
				map.put("DCY", jTextField6.getText());
				map.put("DCRQ", jTextField7.getText());

				int flag = ExcelUtil.transFile(fName, map);
				if (1 == flag) {
					resultJLabel.setForeground(Color.green);
					resultJLabel.setText("转换成功！");
				} else {
					resultJLabel.setForeground(Color.red);
					resultJLabel.setText("转换失败！");
				}
			} catch (Exception e) {
				e.printStackTrace();
				resultJLabel.setForeground(Color.red);
				resultJLabel.setText("转换失败！");
			}

		}
	}
}

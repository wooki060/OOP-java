import java.awt.BorderLayout;
import java.awt.Color;
import java.awt.Component;
import java.awt.Dimension;
import java.awt.Font;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.MouseEvent;
import java.awt.event.MouseListener;
import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.util.Scanner;
import java.util.StringTokenizer;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JMenu;
import javax.swing.JMenuBar;
import javax.swing.JMenuItem;
import javax.swing.JOptionPane;
import javax.swing.JScrollPane;
import javax.swing.JTable;
import javax.swing.JViewport;
import javax.swing.SwingConstants;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.table.DefaultTableCellRenderer;
import javax.swing.table.DefaultTableModel;

public class ExcelDemo extends JFrame{

	private JScrollPane scrollPane;
	private JTable table;
	private JTable headerTable;
	private JMenuBar menuBar;
	private JMenu fileMenu,formulasMenu,funtionMenu;
	private JMenuItem newItem,open,save,exit,sum,average,count,max,min;
	private String title;
	private int cardinality,degree;

	public ExcelDemo(){
		title = "새 Microsoft Excel 워크시트.xlsx - Excel"; //기본 타이틀 
		//창 생성 및 기본 설정 
		setTitle(title);
		setSize(610,488);
		setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		setLocationRelativeTo(null);
		createMenu();

		Object[] colHeader = {"A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z"}; //column헤더 
		Object[][] cel = new String[100][26]; //table 값 
		for(int i=0;i<100;i++)
			for(int j=0;j<26;j++)
				cel[i][j] = "";
		DefaultTableModel cell = new DefaultTableModel(cel,colHeader) { //table 셀 수정 가능
			@Override
			public boolean isCellEditable(int row, int col) {
				return true;
			}
		};
		table = new JTable(cell); //테이블 생성
		table.setAutoResizeMode(JTable.AUTO_RESIZE_OFF); //사이즈 변경 불가  
	    table.getTableHeader().setReorderingAllowed(false); // 컬럼들 이동 불가
        table.getTableHeader().setResizingAllowed(false); // 컬럼 크기 조절 불가 
		table.setCellSelectionEnabled(true); //셀하나만 선택 가능
		
		Object[][] rowHeader = new String[100][1]; //row 헤더를 위한 테이블 
		for(int i=0;i<100;i++)
			rowHeader[i][0] = i+"";
		Object[] n = {""};
		DefaultTableModel header = new DefaultTableModel(rowHeader,n) { //header 테이블 모델 생성, header 테이블은 수정 불가
			//셀 수정 시 col=0이면 수정 불가능 하게 함.(header는 수정 불가)
			@Override
			public boolean isCellEditable(int row, int col){
	    		if(col == 0)
	    			return false;
	    		return true;
	    	}
		};
		
		//header 테이블 생성 및 기본설정 
		headerTable = new JTable(header);  
		table.setSelectionModel(headerTable.getSelectionModel());
		headerTable.setCellSelectionEnabled(false);
	    headerTable.setColumnSelectionAllowed(false);
	    headerTable.setCellSelectionEnabled(false);
	    
	    //row 헤더 생성을 위한 JViewport
		JViewport rowHead = new JViewport();
		rowHead.setView(headerTable);
	    rowHead.setPreferredSize(new Dimension(30,0));
	    
	    //셀 가운데 정렬
	    DefaultTableCellRenderer dtcr1 = new DefaultTableCellRenderer();  
	    dtcr1.setHorizontalAlignment(SwingConstants.CENTER);
	    for(int i=0;i<26;i++)
	    	table.getColumnModel().getColumn(i).setCellRenderer(dtcr1);
	    
	    //column 헤더 가운데 정렬, 색상 변경 
	    DefaultTableCellRenderer colHeadRender = (DefaultTableCellRenderer)table.getTableHeader().getDefaultRenderer(); 
	    colHeadRender.setHorizontalAlignment(SwingConstants.CENTER);
	    table.getTableHeader().setDefaultRenderer(colHeadRender);
	    
	    //row 헤더 가운데 정렬, 선택시 색상,글꼴 변경
	    DefaultTableCellRenderer rowHeadRender = new DefaultTableCellRenderer(){ 
	    	@Override
	    	public Component getTableCellRendererComponent(JTable table, Object value, boolean isSelected, boolean hasFocus, int row, int column) {
	    		Component cell = super.getTableCellRendererComponent(table, value, isSelected, hasFocus, row, column);
	    		int[] rows = table.getSelectedRows(); //int 형 배열에 현재 선택된 row의 값들을 받아옴 
	    		boolean ch;
	    		//모든 열(0~99) 까지 돌면서 선택된 배열이면 설정을 바꿔줌(글꼴, 색, 셀 배경색)
	    		for(int i=0;i<100;i++) {
	    			ch = false;
	    			for(int j=0;j<rows.length;j++) {
	    				if(row == rows[j])
	    					ch = true;
	    			}
	    			if(ch == true) {
	    				cell.setFont(new Font("Arial", Font.BOLD, 15));
	    				cell.setForeground(new Color(0xFFFFFF));
	    				cell.setBackground(Color.blue);
	    			}
	    			else {
	    				cell.setFont(headerTable.getFont());
	    				cell.setForeground(headerTable.getForeground());
	    				cell.setBackground(new Color(0xEEEEEE));
	    			}
	    		}
	    		return cell;
	    	}
	    };
	    rowHeadRender.setHorizontalAlignment(SwingConstants.CENTER); //가운데 정렬 
	    headerTable.getColumnModel().getColumn(0).setCellRenderer(rowHeadRender);
	    
	    //셀테두리 설정
	    table.setGridColor(Color.black);   
	    headerTable.setGridColor(Color.black);
	    
	    //JscrollPane 설정 
		scrollPane = new JScrollPane(table);
		scrollPane.setRowHeader(rowHead);
		scrollPane.setSize(new Dimension(10,10));
		add(scrollPane,BorderLayout.CENTER);
	    
		//현재 마우스 클릭위치 필드값에 저장 
		MouseListener mouse = new MouseListener() {
			@Override
			public void mouseClicked(MouseEvent e) {
				cardinality = table.getSelectedRow();
				degree = table.getSelectedColumn();
			}
			@Override
			public void mousePressed(MouseEvent e) {}
			@Override
			public void mouseReleased(MouseEvent e) {}
			@Override
			public void mouseEntered(MouseEvent e) {}
			@Override
			public void mouseExited(MouseEvent e) {}
			
		};
		table.addMouseListener(mouse);
		
		
		//file menu 설정 
	    newItem.addActionListener(new ActionListener() {
			@Override
			public void actionPerformed(ActionEvent e) {
				//타이틀 재설정 
				title = "새 Microsoft Excel 워크시트.xlsx - Excel";
				setTitle(title);
				//테이블 모델인 cell을 초기화시켜 table 초기화 
				cell.setDataVector(cel, colHeader);
			}
		});
	    open.addActionListener(new ActionListener() {
			@Override
			public void actionPerformed(ActionEvent e) {
				//JFileChooser 선언 및 확장자를 제한하기 위한 filter 설정 
				JFileChooser choose = new JFileChooser();
				choose.setMultiSelectionEnabled(false);
				choose.setFileFilter(new FileNameExtensionFilter("open","txt","csv")); //확장자가 txt와 csv인 파일만 열수 있음 
				//파일이 선택되었는지 확인, 취소하거나 선택하지 않을경우 종료 
				if(choose.showOpenDialog(null) != JFileChooser.APPROVE_OPTION) { 
					JOptionPane.showMessageDialog(null, "파일을 선택하지 않았습니다");
					return;
				}
				//선택된 파일을 입력받기 위해 file,fr,br설정
				File file = choose.getSelectedFile();
				FileReader fr = null;
				BufferedReader br = null;
				try {
					fr = new FileReader(file);
					br = new BufferedReader(fr);
				} catch (FileNotFoundException e1) {
					JOptionPane.showMessageDialog(null, "파일이 존재하지 않습니다");
					e1.printStackTrace();
					return;
				} catch (IOException e1) {
					e1.printStackTrace();
					return;
				}
				//title 재설정 및 cell초기화 
				title = file.toString();
				setTitle(title);
				cell.setDataVector(cel, colHeader);
				//파일 입력으로 table cell 채우기 
				String input=null,val=null;
				int row,column;
				row=0;
				try {
					input = br.readLine(); //한 줄씩 읽음 
				} catch (IOException e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
				}
				while(input != null){
					StringTokenizer st = new StringTokenizer(input,","); //","로 구분해서 한 셀에 하나씩 입력 
					column=0;
					if(input.charAt(0) == ',') cell.setValueAt("", row, column++); //첫 글자가 ','일 경우 첫 셀은 공백으로 그 뒤 부터 stringtoknizer
					while(st.hasMoreTokens()) { //마지막 토큰까지 
						val = st.nextToken();
						cell.setValueAt(val, row, column);  //cell 값 설정 
						column++;
					}
					row++;
					try {
						input = br.readLine(); //다음줄 읽음 
					} catch (IOException e1) {
						// TODO Auto-generated catch block
						e1.printStackTrace();
					}
				}
				try { //종료 
					br.close();
					fr.close();
				} catch (IOException e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
				}
			}
	    });
	    save.addActionListener(new ActionListener() {
			@Override
			public void actionPerformed(ActionEvent e) {
				//JFileChooser 선언 및 초기설정 
				JFileChooser choose = new JFileChooser();
				choose.setMultiSelectionEnabled(false);
				choose.setSelectedFile(new File(title)); //현재 title을 기본이름으로 지정 
				if(choose.showSaveDialog(null) == choose.SAVE_DIALOG) { //저장을 취소할 경우 
					JOptionPane.showMessageDialog(null, "저장 취소");
					return;
				}
				//파일 경로를 string으로 가져와서 확장자명 확인 
				String filePath = choose.getSelectedFile().toString(); 
				StringTokenizer st = new StringTokenizer(filePath,".");
				String filter=null;
				while(st.hasMoreTokens()) { //마지막 '.'까지 옮김 
					filter = st.nextToken();
				}
				//확장자 명이 txt나 csv가 아닐 경우 뒤에 ".txt"를 추가해줌  
				if(!filter.equals("txt") && !filter.equals("csv")) {
					filePath+=".txt";
				}
				//저장할 파일 설정 
				File file = new File(filePath);
				//같은 이름이 있을 경우 덮어쓸지 확인 
				if(file.exists()) {
					if(JOptionPane.showConfirmDialog(null, "동일한 이름의 파일이 존재합니다. 덮어 쓰시겠습니까?","경고",JOptionPane.YES_NO_OPTION) == JOptionPane.NO_OPTION) {
						JOptionPane.showMessageDialog(null, "저장 취소"); //취소할경
						return;
					}
				}
				//타이틀 재설정 
				title = file.toString();
				//출력 파일 열기 
				FileWriter fw = null;
				BufferedWriter bw = null;
				try {
					fw = new FileWriter(file);
					bw = new BufferedWriter(fw);
				} catch (IOException e1) {
					JOptionPane.showMessageDialog(null, "저장 실패");
					e1.printStackTrace();
					return;
				}
				//모든 셀을 돌면서 셀값들을 파일에 출력해줌, 하나 출력마다 ","도 같이 출력 
				int row,column;
				String output=null;
				for(row=0;row<100;row++) {
					for(column=0;column<26;column++) {
						output = table.getValueAt(row, column).toString();
						try {
							bw.write(output);
							if(column != 25) bw.write(", ");  //마지막 column뒤에는 출력하지 않음 
						} catch (IOException e1) {
							e1.printStackTrace();
						}
					}
					try {
						bw.newLine(); //줄바꿈 
					} catch (IOException e1) {
						e1.printStackTrace();
					}
				}
				try { //파일 종료 
					bw.close();
					fw.close();
				} catch (IOException e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
				}
				setTitle(title);
			}
		});
	    exit.addActionListener(new ActionListener() {
	    	@Override
	    	public void actionPerformed(ActionEvent e) {
	    		// TODO Auto-generated method stub
				System.exit(0); //종료 
	    	}
	    });
	    
	    //formular menu 설정 
	    sum.addActionListener(new ActionListener() {  	    	
	    	@Override
	    	public void actionPerformed(ActionEvent e) {
	    		//input dialog 생성, range에 반환값 저장.
	    		String range = JOptionPane.showInputDialog(null, "Function Arguments.", "SUM", JOptionPane.YES_NO_OPTION);
	    		if(range == null || range.indexOf(":") == -1) { //range가 null이면 입력이 없음, 입력에 :가 존재하지 않으면 잘못된 입력 
	    			JOptionPane.showMessageDialog(null, "A0:Z99형식으로 입력하세요");
	    		}else {
	    			StringTokenizer st = null;
	    			String start=null,end=null;
	    			st = new StringTokenizer(range,":"); //:를 기준으로 나눔 
	    			start = st.nextToken(); //시작지점 
	    			end = st.nextToken(); //종료지점 
	    			
	    			int sx=0,sy=0,ex=0,ey=0;
	    			start.toUpperCase();
	    			sy = (int)start.charAt(0) - 'A'; //시작지점의 x좌표 
	    			sx = Integer.parseInt(start.substring(1)); //시작지점의 y좌표 
	    			end.toUpperCase();
	    			ey = (int)end.charAt(0) - 'A';  //종료지점의 x좌표 
	    			ex = Integer.parseInt(end.substring(1)); //종료지점의 y좌표 
	    			
	    			double sum=0;
	    			for(int i=sx;i<=ex;i++) //range의 시작부터 끝까지 반복 
	    				for(int j=sy;j<=ey;j++) {
	    					String ce = cell.getValueAt(i, j).toString(); //string으로 cell 값을 받아옴 
	    					Scanner sc = new Scanner(ce.trim());
	    					if(!ce.isEmpty() && sc.hasNextDouble()) //빈 셀이거나 숫자가 아닐 경우 무시 
	    						sum += Double.parseDouble(ce);
	    				}
	    			cell.setValueAt(String.valueOf(sum), cardinality, degree); //sum값 double로 cell에 입력 
	    		}
	    	}
	    });
	    //sum과 비슷한 방식으로 나머지 action들도 처리 
	    average.addActionListener(new ActionListener() {
	    	@Override
	    	public void actionPerformed(ActionEvent e) {
	    		String range = JOptionPane.showInputDialog(null, "Function Arguments.", "AVERAGE", JOptionPane.YES_NO_OPTION);
	    		if(range == null || range.indexOf(":") == -1) {
	    			JOptionPane.showMessageDialog(null, "A0:Z99형식으로 입력하세요");
	    		}else {
	    			StringTokenizer st = null;
	    			String start=null,end=null;
	    			st = new StringTokenizer(range,":");
	    			start = st.nextToken();
	    			end = st.nextToken();
	    			
	    			int sx=0,sy=0,ex=0,ey=0;
	    			start.toUpperCase();
	    			sy = (int)start.charAt(0) - 'A';
	    			sx = Integer.parseInt(start.substring(1));
	    			end.toUpperCase();
	    			ey = (int)end.charAt(0) - 'A';
	    			ex = Integer.parseInt(end.substring(1));
	    			
	    			double sum=0;
	    			int cnt=0;
	    			for(int i=sx;i<=ex;i++)
	    				for(int j=sy;j<=ey;j++) {
	    					String ce = cell.getValueAt(i, j).toString();
	    					Scanner sc = new Scanner(ce.trim());
	    					if(!ce.isEmpty() && sc.hasNextDouble()) {
	    						sum += Double.parseDouble(ce);
	    						cnt++; //평균을 구하기 위해 갯수도 세줌 	    			
	    					}
	    				}	
	    			if(cnt != 0) cell.setValueAt(String.valueOf(sum/cnt), cardinality, degree);
	    			else {
	    				JOptionPane.showMessageDialog(null, "0으로 나눌 수 없습니다.");
	    			}
	    		}
	    	}
	    });
	    count.addActionListener(new ActionListener() {
	    	@Override
	    	public void actionPerformed(ActionEvent e) {
	    		// TODO Auto-generated method stub
	    		String range = JOptionPane.showInputDialog(null, "Function Arguments.", "COUNT", JOptionPane.YES_NO_OPTION);
	    		if(range == null || range.indexOf(":") == -1) {
	    			JOptionPane.showMessageDialog(null, "A0:Z99형식으로 입력하세요");
	    		}else {
	    			StringTokenizer st = null;
	    			String start=null,end=null;
	    			st = new StringTokenizer(range,":");
	    			start = st.nextToken();
	    			end = st.nextToken();
	    			
	    			int sx=0,sy=0,ex=0,ey=0;
	    			start.toUpperCase();
	    			sy = (int)start.charAt(0) - 'A';
	    			sx = Integer.parseInt(start.substring(1));
	    			end.toUpperCase();
	    			ey = (int)end.charAt(0) - 'A';
	    			ex = Integer.parseInt(end.substring(1));
	    		
	    			int cnt=0;
	    			for(int i=sx;i<=ex;i++)
	    				for(int j=sy;j<=ey;j++) {
	    					String ce = cell.getValueAt(i, j).toString();
	    					Scanner sc = new Scanner(ce.trim());
	    					if(!ce.isEmpty() && sc.hasNextDouble()) { 
	    						cnt++;		
	    					}
	    				}
	    			cell.setValueAt(String.valueOf(cnt), cardinality, degree);
	    		}
	    	}
	    });
	    max.addActionListener(new ActionListener() {
	    	@Override
	    	public void actionPerformed(ActionEvent e) {
	    		// TODO Auto-generated method stub
	    		String range = JOptionPane.showInputDialog(null, "Function Arguments.", "MAX", JOptionPane.YES_NO_OPTION);
	    		if(range == null || range.indexOf(":") == -1) {
	    			JOptionPane.showMessageDialog(null, "A0:Z99형식으로 입력하세요");
	    		}else {
	    			StringTokenizer st = null;
	    			String start=null,end=null;
	    			st = new StringTokenizer(range,":");
	    			start = st.nextToken();
	    			end = st.nextToken();
	    		
	    			int sx=0,sy=0,ex=0,ey=0;
	    			start.toUpperCase();
	    			sy = (int)start.charAt(0) - 'A';
	    			sx = Integer.parseInt(start.substring(1));
	    			end.toUpperCase();
	    			ey = (int)end.charAt(0) - 'A';
	    			ex = Integer.parseInt(end.substring(1));
	    			
	    			double max=Double.MIN_VALUE; //최대값을 구하기 위해 최소값으로 초기화 
	    			for(int i=sx;i<=ex;i++)
	    				for(int j=sy;j<=ey;j++) {
	    					String ce = cell.getValueAt(i, j).toString();
	    					Scanner sc = new Scanner(ce.trim());
	    					//빈 셀이거나 숫자가 아닌 경우를 제외하고 현재까지 max값보다 크면 max 값 바꿈
	    					if(!ce.isEmpty() && sc.hasNextDouble() && max < Double.parseDouble(ce)) { 
	    						max = Double.parseDouble(ce);
	    					}
	    				}
	    			cell.setValueAt(String.valueOf(max), cardinality, degree);
	    		}
	    	}
	    });
	    min.addActionListener(new ActionListener() {
	    	@Override
	    	public void actionPerformed(ActionEvent e) {
	    		// TODO Auto-generated method stub
	    		String range = JOptionPane.showInputDialog(null, "Function Arguments.", "MIN", JOptionPane.YES_NO_OPTION);
	    		if(range == null || range.indexOf(":") == -1) {
	    			JOptionPane.showMessageDialog(null, "A0:Z99형식으로 입력하세요");
	    		}else {
	    			StringTokenizer st = null;
	    			String start=null,end=null;
	    			st = new StringTokenizer(range,":");
	    			start = st.nextToken();
	    			end = st.nextToken();
	    			
	    			int sx=0,sy=0,ex=0,ey=0;
	    			start.toUpperCase();
	    			sy = (int)start.charAt(0) - 'A';
	    			sx = Integer.parseInt(start.substring(1));
	    			end.toUpperCase();
	    			ey = (int)end.charAt(0) - 'A';
	    			ex = Integer.parseInt(end.substring(1));
	    		
	    			double min=Double.MAX_VALUE;
	    			for(int i=sx;i<=ex;i++)
	    				for(int j=sy;j<=ey;j++) {
	    					String ce = cell.getValueAt(i, j).toString();
	    					Scanner sc = new Scanner(ce.trim());
	    					//max와 반대로 현재까지 최소값보다 작으면 min값 바꿈 
	    					if(!ce.isEmpty() && sc.hasNextDouble() && min > Double.parseDouble(ce)) {
	    						min = Double.parseDouble(ce);
	    						
	    					}
	    				}
	    			cell.setValueAt(String.valueOf(min), cardinality, degree);
	    		}
	    	}
	    });
		
		setVisible(true);
	}
	//메뉴 생성 
	public void createMenu(){
		menuBar = new JMenuBar();
		//메뉴바에 file과 formulas메뉴를 추가 
		fileMenu = new JMenu("File");
		formulasMenu = new JMenu("Formulas");
		menuBar.add(fileMenu);
		menuBar.add(formulasMenu);
		//file 메뉴에 new,open,save,exit 추가 
		newItem = new JMenuItem("New"); 
		open = new JMenuItem("Open");
		save = new JMenuItem("Save");
		exit = new JMenuItem("Exit");
		fileMenu.add(newItem);
		fileMenu.add(open);
		fileMenu.add(save);
		fileMenu.add(exit);
		//formulas 메뉴에 sum,average,count,max,min추가 
		sum = new JMenuItem("SUM");
		average = new JMenuItem("AVERAGE");
		count = new JMenuItem("COUNT");
		max = new JMenuItem("MAX");
		min = new JMenuItem("MIN");
		formulasMenu.add(sum);
		formulasMenu.add(average);
		formulasMenu.add(count);
		formulasMenu.add(max);
		formulasMenu.add(min);
		//메뉴바 시작 
		setJMenuBar(menuBar);
	}
	//데모파일 창 시작 
	public static void main(String[] args) {
		ExcelDemo excel = new ExcelDemo();
	}
}

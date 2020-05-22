package Assignment1;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.PrintWriter;
import java.util.Calendar;
import java.util.Scanner;

public class TravelInfoRequest {
	
	public static void main(String[] args) {
		// TODO Auto-generated method stub
		Scanner properties_input = null;
		Scanner template_input = null;
		PrintWriter output = null;
		try {
			properties_input = new Scanner(new FileInputStream("properties.txt"));
			template_input = new Scanner(new FileInputStream("template_file.txt"));
			output = new PrintWriter("output_file.txt");
			final int max_size = 1000; // properties 파일의 최대 KeyValue 개수는 1000 으로 한다

			Distance stDis = null,deDis = null;
			Countries start=null,depart=null;
			int count;
			KeyValue[] properties = new KeyValue[max_size];
			Calendar cal = Calendar.getInstance();
			properties[0] = new KeyValue("date",cal.get(Calendar.YEAR)+"-"+(cal.get(Calendar.MONTH)+1)+"-"+cal.get(Calendar.DATE));
			count=1;
			while(properties_input.hasNextLine()) {
				properties[count] = new KeyValue(properties_input.nextLine());
				
				if(properties[count].getKey().equals("startcountry")) {
					start = new Countries(properties[count].getValue());
					stDis = new Distance(start.getCountry(),start.getLat(),start.getLon());
				}
				else if(properties[count].getKey().equals("departcountry")) {
					depart = new Countries(properties[count].getValue());
					deDis = new Distance(depart.getCountry(),depart.getLat(),depart.getLon());
				}
				
				count++;
			}
			
			String get_line, now_key;
			while(template_input.hasNextLine()) {
				get_line = template_input.nextLine();
				if(get_line.equals("<add info>")) {
					get_line = stDis.getDistance(stDis, deDis);
				}
				else {
					for(int i = 0; i < count; i++) {
						now_key = "{" + properties[i].getKey() + "}";
						get_line = get_line.replace(now_key,properties[i].getValue());
					}
				}
				output.println(get_line);
			}
			
		} catch (FileNotFoundException e) {
			System.out.println("File not found.");
			System.exit(0);
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			try {
				if (properties_input != null) properties_input.close();
				if (template_input != null) template_input.close();
				if (output != null) output.close();
			} catch (Exception e2) {
				e2.printStackTrace();
			}
		}
	}

}
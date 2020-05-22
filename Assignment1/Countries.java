package Assignment1;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.util.Scanner;
import java.util.StringTokenizer;

public class Countries {
	private String country;
	private double lat, lng;

	public Countries(String country) {
		Scanner countries_input = null;
		try {
			countries_input = new Scanner(new FileInputStream("Countries.csv"));
			final int country_size = 245;
			String[] countries = new String[country_size];
			
			for(int i=0; i<country_size;i++) {
				countries[i] = countries_input.nextLine();
				
				StringTokenizer str = new StringTokenizer(countries[i],",");
				if(country.equals(str.nextToken())) {
					this.country = country;
					this.lat = Double.parseDouble(str.nextToken());
					this.lng = Double.parseDouble(str.nextToken());
				}
			}
			
		} catch (FileNotFoundException e) {
			System.out.println("File not found.");
			System.exit(0);
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			try {
				if (countries_input != null) countries_input.close();
			} catch (Exception e2) {
				e2.printStackTrace();
			}
		}
	}
	
	public String getCountry() {return country;}
	public double getLat() {return lat;}
	public double getLon() {return lng;}
}
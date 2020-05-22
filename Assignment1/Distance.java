package Assignment1;

public class Distance {
	private String name;
	private double lat, lng;
	
	public Distance(String name,double lat, double lng) {
		this.name = name;
		this.lat  = lat;
		this.lng = lng;
	}
	public String writeDistance() {
		String info="Country : "+name+"\nlatitude = "+lat+"\nlongitude = "+lng+"\n------------------------\n";
		return info;
	}
	public String getDistance(Distance a,Distance b) {
		double distance=Math.sqrt(Math.pow(a.lat-b.lat, 2) + Math.pow(a.lng-b.lng, 2));
		String addInfo="Distance of\n"+a.writeDistance()+b.writeDistance()+"is\n" + distance;
		return addInfo;
	}
}

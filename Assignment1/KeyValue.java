package Assignment1;

import java.util.StringTokenizer;

public class KeyValue {
	
	private String key;
	private String value;
	
	public KeyValue(String kv) {
		StringTokenizer str = new StringTokenizer(kv,"=");
		this.key = str.nextToken();
		this.value = str.nextToken();
		while(str.hasMoreTokens())
			this.value += str.nextToken();
	}
	
	public KeyValue(String key,String value) {
		this.key = key;
		this.value = value;
	}
	public String getKey() {return key;}
	public String getValue() {return value;}
	
}
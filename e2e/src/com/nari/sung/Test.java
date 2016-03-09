package com.nari.sung;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

public class Test {

	/**
	 * @param args
	 */
	public static void main(String[] args) {
		// System.out.println(zeroPadder("12", 4));
		String zm = "02";

		if (zm.indexOf("组") != -1) {
			zm = zm.substring(0, zm.indexOf("组"));
		}
		zm = zeroPadder(zm, 2);

		System.out.println(zm);
	}

	public static String zeroPadder(String s, int order) {
		if (s.length() >= order) {
			return s;
		}
		char[] data = new char[order];
		Arrays.fill(data, '0');
		for (int i = s.length() - 1, j = order - 1; i >= 0; i--, j--) {
			data[j] = s.charAt(i);
		}
		return String.valueOf(data);
	}
}

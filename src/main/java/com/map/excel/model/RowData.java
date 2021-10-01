/**
 * Code created by Rohit Bhatia for self use or Demo purpose only.
 */
package com.map.excel.model;

import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.Set;

/**
 * @author blchi
 *
 */
public class RowData {
	String txCode = "";
	BigDecimal bgSumTxCount = new BigDecimal(0);
	BigDecimal bgSumStepCount = new BigDecimal(0);
	BigDecimal bgSumGuiTime = new BigDecimal(0);
	private List<String> roles = new ArrayList<String>();
	//private List<String> users = new ArrayList<String>();
	private Set<String> users = new HashSet<>();
	//private List<String> uniqueUsers = new ArrayList<String>();
	public String getTxCode() {
		return txCode;
	}

	public void setTxCode(String txCode) {
		this.txCode = txCode;
	}

	public BigDecimal getBgSumTxCount() {
		return bgSumTxCount;
	}

	public void setBgSumTxCount(BigDecimal bgSumTxCount) {
		this.bgSumTxCount = bgSumTxCount;
	}

	public BigDecimal getBgSumStepCount() {
		return bgSumStepCount;
	}

	public void setBgSumStepCount(BigDecimal bgSumStepCount) {
		this.bgSumStepCount = bgSumStepCount;
	}

	public BigDecimal getBgSumGuiTime() {
		return bgSumGuiTime;
	}

	public void setBgSumGuiTime(BigDecimal bgSumGuiTime) {
		this.bgSumGuiTime = bgSumGuiTime;
	}

	public List<String> getRoles() {
		return roles;
	}

	public void setRoles(List<String> roles) {
		this.roles = roles;
	}

	public Set<String> getUsers() {
		return users;
	}

	public void setUsers(Set<String> users) {
		this.users = users;
	}

	/*
	 * public List<String> getUsers() { return users; }
	 * 
	 * public void setUsers(List<String> users) { this.users = users; }
	 */

	/*
	 * public List<String> getUniqueUsers() { if (!users.isEmpty()) this.uniqueUsers
	 * = users.stream().distinct().collect(Collectors.toList()); return uniqueUsers;
	 * }
	 */

}

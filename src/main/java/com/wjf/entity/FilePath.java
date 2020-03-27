package com.wjf.entity;

import java.io.Serializable;

/**
 * 文件路径
 */
public class FilePath implements Serializable {

	private static final long serialVersionUID = 5861683791118940868L;

	//	真实路径
	private String realPath;

//	用于映射的路径
	private String handlerPath;

	public String getRealPath() {
		return realPath;
	}

	public void setRealPath(String realPath) {
		this.realPath = realPath;
	}

	public String getHandlerPath() {
		return handlerPath;
	}

	public void setHandlerPath(String handlerPath) {
		this.handlerPath = handlerPath;
	}

	public FilePath() {
	}

	public FilePath(String realPath, String handlerPath) {
		this.realPath = realPath;
		this.handlerPath = handlerPath;
	}
}

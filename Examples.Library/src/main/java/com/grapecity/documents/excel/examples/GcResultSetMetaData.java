package com.grapecity.documents.excel.examples;

import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.util.EnumSet;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.UsedRangeType;

public class GcResultSetMetaData implements ResultSetMetaData {
	private IWorksheet worksheet = null;
	
	public GcResultSetMetaData(IWorksheet worksheet) {
		this.worksheet = worksheet;
	}

	@Override
	public <T> T unwrap(Class<T> paramClass) throws SQLException {
		// TODO Auto-generated method stub
		return null;
	}

	@Override
	public boolean isWrapperFor(Class<?> paramClass) throws SQLException {
		// TODO Auto-generated method stub
		return false;
	}

	@Override
	public int getColumnCount() throws SQLException {
		return this.worksheet.getUsedRange(EnumSet.of(UsedRangeType.Data)).getColumnCount();
	}

	@Override
	public boolean isAutoIncrement(int paramInt) throws SQLException {
		// TODO Auto-generated method stub
		return false;
	}

	@Override
	public boolean isCaseSensitive(int paramInt) throws SQLException {
		// TODO Auto-generated method stub
		return false;
	}

	@Override
	public boolean isSearchable(int paramInt) throws SQLException {
		// TODO Auto-generated method stub
		return false;
	}

	@Override
	public boolean isCurrency(int paramInt) throws SQLException {
		// TODO Auto-generated method stub
		return false;
	}

	@Override
	public int isNullable(int paramInt) throws SQLException {
		// TODO Auto-generated method stub
		return 0;
	}

	@Override
	public boolean isSigned(int paramInt) throws SQLException {
		// TODO Auto-generated method stub
		return false;
	}

	@Override
	public int getColumnDisplaySize(int paramInt) throws SQLException {
		// TODO Auto-generated method stub
		return 0;
	}

	@Override
	public String getColumnLabel(int paramInt) throws SQLException {
		Object columnLabel = this.worksheet.getRange(0, paramInt - 1).getValue();
		
		if (columnLabel == null) {
			return null;
		}
		else {
			return columnLabel.toString();
		}
	}

	@Override
	public String getColumnName(int paramInt) throws SQLException {
		// TODO Auto-generated method stub
		return null;
	}

	@Override
	public String getSchemaName(int paramInt) throws SQLException {
		// TODO Auto-generated method stub
		return null;
	}

	@Override
	public int getPrecision(int paramInt) throws SQLException {
		// TODO Auto-generated method stub
		return 0;
	}

	@Override
	public int getScale(int paramInt) throws SQLException {
		// TODO Auto-generated method stub
		return 0;
	}

	@Override
	public String getTableName(int paramInt) throws SQLException {
		// TODO Auto-generated method stub
		return null;
	}

	@Override
	public String getCatalogName(int paramInt) throws SQLException {
		// TODO Auto-generated method stub
		return null;
	}

	@Override
	public int getColumnType(int paramInt) throws SQLException {
		// TODO Auto-generated method stub
		return 0;
	}

	@Override
	public String getColumnTypeName(int paramInt) throws SQLException {
		// TODO Auto-generated method stub
		return null;
	}

	@Override
	public boolean isReadOnly(int paramInt) throws SQLException {
		// TODO Auto-generated method stub
		return false;
	}

	@Override
	public boolean isWritable(int paramInt) throws SQLException {
		// TODO Auto-generated method stub
		return false;
	}

	@Override
	public boolean isDefinitelyWritable(int paramInt) throws SQLException {
		// TODO Auto-generated method stub
		return false;
	}

	@Override
	public String getColumnClassName(int paramInt) throws SQLException {
		// TODO Auto-generated method stub
		return null;
	}

}

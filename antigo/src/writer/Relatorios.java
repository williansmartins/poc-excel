package writer;

import java.io.File;
import java.io.IOException;
import java.util.Locale;

import jxl.CellView;
import jxl.Workbook;
import jxl.WorkbookSettings;
import jxl.format.UnderlineStyle;
import jxl.write.Formula;
import jxl.write.Label;
import jxl.write.Number;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;


public class WriteExcel {

	public void generateReportWithDetails(){
		
		HSSFWorkbook wb = new HSSFWorkbook();
		HSSFSheet sheet = wb.createSheet();
		HSSFRow row;
		
		row = sheet.createRow(0);
		row.createCell(0).setCellValue("Nome da Franquia");
		row.createCell(1).setCellValue("CNPJ");
		row.createCell(2).setCellValue("Colaborador"); 
		row.createCell(3).setCellValue("Cargo"); 
		
		personsFranquias = service.findAllFranquiaAndUsers();
		
		for (int i = 1; i <= personsFranquias.size(); i++){
			PersonFranquiaVO vo = personsFranquias.get(i-1);
			row = sheet.createRow(i);
			row.createCell(0).setCellValue(vo.getFranquia());
			row.createCell(1).setCellValue(vo.getCnpj()); 
			row.createCell(2).setCellValue(vo.getColaborador()); 
			row.createCell(3).setCellValue(vo.getCargo()); 
		}
		
		//formatar a tabela
		sheet.autoSizeColumn(0);
		sheet.autoSizeColumn(1);
		sheet.autoSizeColumn(2);
		sheet.autoSizeColumn(3);
		
		HttpServletResponse res = (HttpServletResponse)FacesContext.getCurrentInstance().getExternalContext().getResponse();
		res.setContentType("application/vnd.ms-excel");
		res.setHeader("Content-disposition",  "attachment; filename=franquias-e-colaboradores.xls");
		
		try {
			ServletOutputStream out = res.getOutputStream();
			
			wb.write(out);
			out.flush();
			out.close();
		} catch (IOException ex) { 
			ex.printStackTrace();
		}
		
		FacesContext faces = FacesContext.getCurrentInstance();
		faces.responseComplete();  
	}
	
} 
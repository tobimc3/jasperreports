package promis.jasper;

import java.io.File;
import java.util.Locale;
import java.util.ResourceBundle;

import net.sf.jasperreports.engine.JRException;
import net.sf.jasperreports.engine.JasperPrint;
import net.sf.jasperreports.engine.export.ooxml.JRXlsxExporter;
import net.sf.jasperreports.export.SimpleExporterInput;
import net.sf.jasperreports.export.SimpleOutputStreamExporterOutput;
import net.sf.jasperreports.export.SimpleXlsxReportConfiguration;
import net.sf.jasperreports.view.JRSaveContributor;

/*********************************
 * Implementation of a XLSXSaveContributor, because there was none.
 * Yay-I rule!
 *
 *
 */




public class XLSXSaveContributor extends JRSaveContributor
{
	
	private Locale locale;
	private ResourceBundle bundle;
	
	public XLSXSaveContributor(Locale locale, ResourceBundle resBundle){
	   this.locale = locale;
	   this.bundle = resBundle;
	}
	
	
	
	

	@Override
	public void save(JasperPrint jp, File pathToFile) throws JRException {
		JRXlsxExporter exporter = new JRXlsxExporter();
		exporter.setExporterInput(new SimpleExporterInput(jp));
	    exporter.setExporterOutput(new SimpleOutputStreamExporterOutput(pathToFile + ".xlsx"));
	    SimpleXlsxReportConfiguration configuration = new SimpleXlsxReportConfiguration(); 
	    exporter.setConfiguration(configuration);
	    exporter.exportReport();

	}

	@Override
	public boolean accept(File f) {
	  return f.getName().endsWith(".xlsx");			
	}

	@Override
	public String getDescription() {
		return "XLSX (*.xlsx)";
	}

}

package com.CrearEvidencias.evidencias;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Scanner;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;

@SpringBootApplication
public class EvidenciasApplication {

	public static void main(String[] args) {
		SpringApplication.run(EvidenciasApplication.class, args);
		String excelPath = "src/main/resources/pasos.xlsx";
		String wordTemplate = "src/main/resources/evidenciabase.docx";
		String outputDir = "src/main/resources/";
		try {
			String iniciativa = leerCeldaExcel(excelPath, 7, 0); // H1 (col 7, row 0)
			String proyecto = leerCeldaExcel(excelPath, 8, 0);   // I1 (col 8, row 0)
			List<Caso> casos = leerCasosDesdeExcelAvanzado(excelPath);
			String fechaHoy = LocalDate.now().format(DateTimeFormatter.ofPattern("dd/MM/yyyy"));
			for (int i = 0; i < casos.size(); i++) {
				String nombreCaso = (casos.get(i).nombre != null && !casos.get(i).nombre.isBlank()) ? casos.get(i).nombre : ("caso_" + (i+1));
				nombreCaso = nombreCaso.replaceAll("[\\\\/:*?\"<>|]", "_");
				String outputWord = outputDir + nombreCaso + ".docx";
				// Obtiene solo el nombre del archivo sin ruta ni extensión, compatible con cualquier separador
				String nombreArchivo = outputWord.replaceAll("^.*[\\\\/](.*)\\.docx$", "$1");
				String nombreParaSelector = nombreArchivo;
				String prerequisitos = casos.get(i).prerequisitos;
				generarWordDesdePlantilla(wordTemplate, outputWord, casos.get(i).pasos, casos.get(i).resultados, nombreParaSelector, prerequisitos, fechaHoy, iniciativa, proyecto);
				System.out.println("Generado: " + outputWord);
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	// Lee una celda específica del Excel (columna, fila)
	public static String leerCeldaExcel(String excelPath, int col, int row) throws IOException {
		try (FileInputStream fis = new FileInputStream(excelPath);
			Workbook workbook = new XSSFWorkbook(fis)) {
			Sheet sheet = workbook.getSheetAt(0);
			Row r = sheet.getRow(row);
			if (r != null) {
				Cell c = r.getCell(col);
				if (c != null) {
					c.setCellType(CellType.STRING);
					return c.getStringCellValue();
				}
			}
		}
		return "";
	}

	// Reemplaza marcadores {pasoX}, {ResultadoEsperadoX}, {nombre}, {prerequisitos}, {fecha}, {iniciativa}, {proyecto}
	public static void reemplazarMarcadoresEnParrafo(XWPFParagraph p, List<String> pasos, List<String> resultados, String nombreCaso, String prerequisitos, String fecha, String iniciativa, String proyecto) {
		StringBuilder fullText = new StringBuilder();
		List<XWPFRun> runs = p.getRuns();
		for (XWPFRun run : runs) {
			String text = run.getText(0);
			if (text != null) fullText.append(text);
		}
		String textoParrafo = fullText.toString();
		boolean modificado = false;
		for (int i = 0; i < pasos.size(); i++) {
			String marcadorPaso = "{paso" + (i+1) + "}";
			if (textoParrafo.contains(marcadorPaso)) {
				textoParrafo = textoParrafo.replace(marcadorPaso, pasos.get(i));
				modificado = true;
			}
			if (resultados != null && resultados.size() > i) {
				String marcadorResultado = "{ResultadoEsperado" + (i+1) + "}";
				if (textoParrafo.contains(marcadorResultado)) {
					textoParrafo = textoParrafo.replace(marcadorResultado, resultados.get(i));
					modificado = true;
				}
			}
		}
		if (nombreCaso != null && nombreCaso.trim().length() > 1 && textoParrafo.contains("{nombre}")) {
			textoParrafo = textoParrafo.replace("{nombre}", nombreCaso);
			modificado = true;
		}
		if (prerequisitos != null && prerequisitos.trim().length() > 1 && textoParrafo.contains("{prerequisitos}")) {
			textoParrafo = textoParrafo.replace("{prerequisitos}", prerequisitos);
			modificado = true;
		}
		if (fecha != null && textoParrafo.contains("{fecha}")) {
			textoParrafo = textoParrafo.replace("{fecha}", fecha);
			modificado = true;
		}
		if (iniciativa != null && textoParrafo.contains("{iniciativa}")) {
			textoParrafo = textoParrafo.replace("{iniciativa}", iniciativa);
			modificado = true;
		}
		if (proyecto != null && textoParrafo.contains("{proyecto}")) {
			textoParrafo = textoParrafo.replace("{proyecto}", proyecto);
			modificado = true;
		}
		if (modificado) {
			for (int i = runs.size() - 1; i >= 0; i--) {
				p.removeRun(i);
			}
			XWPFRun newRun = p.createRun();
			newRun.setText(textoParrafo);
		}
	}

	// Lee los pasos, resultados, nombre y prerequisitos desde un archivo Excel y agrupa por caso
	public static List<Caso> leerCasosDesdeExcelAvanzado(String excelPath) throws IOException {
		List<Caso> casos = new ArrayList<>();
		List<String> pasosActual = new ArrayList<>();
		List<String> resultadosActual = new ArrayList<>();
		String nombreCasoActual = null;
		String prerequisitosActual = null;
		try (FileInputStream fis = new FileInputStream(excelPath);
			Workbook workbook = new XSSFWorkbook(fis)) {
			Sheet sheet = workbook.getSheetAt(0);
			for (Row row : sheet) {
				Cell cellNumero = row.getCell(0); // Columna 0: número de paso
				Cell cellDescripcion = row.getCell(1); // Columna 1: descripción
				Cell cellResultado = row.getCell(2); // Columna 2: resultado esperado
				Cell cellNombre = row.getCell(5); // Columna F (índice 5): nombre del caso
				Cell cellPrereq = row.getCell(6); // Columna G (índice 6): prerequisitos
				if (cellNumero == null || cellDescripcion == null) continue;
				int numeroPaso = (int) cellNumero.getNumericCellValue();
				String descripcion = cellDescripcion.getStringCellValue();
				String resultado = (cellResultado != null) ? cellResultado.getStringCellValue() : "";
				if (numeroPaso == 1 && !pasosActual.isEmpty()) {
					casos.add(new Caso(new ArrayList<>(pasosActual), new ArrayList<>(resultadosActual), nombreCasoActual, prerequisitosActual));
					pasosActual.clear();
					resultadosActual.clear();
					nombreCasoActual = null;
					prerequisitosActual = null;
				}
				pasosActual.add(descripcion);
				resultadosActual.add(resultado);
				if (numeroPaso == 1 && cellNombre != null) {
					nombreCasoActual = cellNombre.getStringCellValue();
				}
				if (numeroPaso == 1 && cellPrereq != null) {
					prerequisitosActual = cellPrereq.getStringCellValue();
				}
			}
			if (!pasosActual.isEmpty()) {
				casos.add(new Caso(pasosActual, resultadosActual, nombreCasoActual, prerequisitosActual));
			}
		}
		return casos;
	}

	public static class Caso {
		public List<String> pasos;
		public List<String> resultados;
		public String nombre;
		public String prerequisitos;
		public Caso(List<String> pasos, List<String> resultados, String nombre, String prerequisitos) {
			this.pasos = pasos;
			this.resultados = resultados;
			this.nombre = nombre;
			this.prerequisitos = prerequisitos;
		}
	}

	// Elimina todo el contenido de las secciones (hojas) siguientes a la del último marcador
	public static void eliminarSeccionesDespuesUltimoMarcador(XWPFDocument doc, List<String> pasos, List<String> resultados) {
		int lastSectionIndex = -1;
		String ultimoMarcador = "";
		if (pasos != null && !pasos.isEmpty()) {
			ultimoMarcador = "{paso" + pasos.size() + "}";
		}
		if (resultados != null && !resultados.isEmpty()) {
			String marcadorResultado = "{ResultadoEsperado" + resultados.size() + "}";
			if (marcadorResultado.compareTo(ultimoMarcador) > 0) {
				ultimoMarcador = marcadorResultado;
			}
		}
		List<IBodyElement> bodyElements = doc.getBodyElements();
		for (int i = 0; i < bodyElements.size(); i++) {
			IBodyElement elem = bodyElements.get(i);
			if (elem instanceof XWPFParagraph) {
				String text = ((XWPFParagraph) elem).getText();
				if (text != null && text.contains(ultimoMarcador)) {
					lastSectionIndex = i;
				}
			}
		}
		// Busca el siguiente salto de sección después del último marcador
		int sectionBreakIndex = -1;
		for (int i = lastSectionIndex + 1; i < bodyElements.size(); i++) {
			IBodyElement elem = bodyElements.get(i);
			if (elem instanceof XWPFParagraph) {
				XWPFParagraph para = (XWPFParagraph) elem;
				if (para.getCTP().isSetPPr() && para.getCTP().getPPr().isSetSectPr()) {
					sectionBreakIndex = i;
					break;
				}
			}
		}
		if (sectionBreakIndex != -1) {
			for (int i = bodyElements.size() - 1; i >= sectionBreakIndex; i--) {
				if (bodyElements.get(i) instanceof XWPFParagraph) {
					doc.removeBodyElement(doc.getPosOfParagraph((XWPFParagraph) bodyElements.get(i)));
				} else if (bodyElements.get(i) instanceof XWPFTable) {
					doc.removeBodyElement(doc.getPosOfTable((XWPFTable) bodyElements.get(i)));
				}
			}
		}
	}

	// Elimina la tabla (o párrafo) que contiene el siguiente marcador ({pasoN+1} o {ResultadoEsperadoN+1}) y todo lo que sigue
	public static void eliminarDesdeSiguienteMarcador(XWPFDocument doc, List<String> pasos, List<String> resultados) {
		String siguienteMarcadorPaso = "{paso" + (pasos.size() + 1) + "}";
		String siguienteMarcadorResultado = "{ResultadoEsperado" + (resultados.size() + 1) + "}";
		List<IBodyElement> bodyElements = doc.getBodyElements();
		int startRemoveIndex = -1;
		boolean existeMarcadorSiguiente = false;
		for (int i = 0; i < bodyElements.size(); i++) {
			IBodyElement elem = bodyElements.get(i);
			if (elem instanceof XWPFTable) {
				XWPFTable table = (XWPFTable) elem;
				outer:
				for (var row : table.getRows()) {
					for (var cell : row.getTableCells()) {
						for (XWPFParagraph p : cell.getParagraphs()) {
							String text = p.getText();
							if ((text != null && text.contains(siguienteMarcadorPaso)) || (text != null && text.contains(siguienteMarcadorResultado))) {
								startRemoveIndex = i;
								existeMarcadorSiguiente = true;
								break outer;
							}
						}
					}
				}
			} else if (elem instanceof XWPFParagraph) {
				String text = ((XWPFParagraph) elem).getText();
				if ((text != null && text.contains(siguienteMarcadorPaso)) || (text != null && text.contains(siguienteMarcadorResultado))) {
					startRemoveIndex = i;
					existeMarcadorSiguiente = true;
					break;
				}
			}
		}
		// Solo elimina si realmente existe un marcador siguiente
		if (existeMarcadorSiguiente && startRemoveIndex != -1) {
			for (int i = bodyElements.size() - 1; i >= startRemoveIndex; i--) {
				IBodyElement elem = bodyElements.get(i);
				if (elem instanceof XWPFParagraph) {
					doc.removeBodyElement(doc.getPosOfParagraph((XWPFParagraph) elem));
				} else if (elem instanceof XWPFTable) {
					doc.removeBodyElement(doc.getPosOfTable((XWPFTable) elem));
				}
			}
		}
	}

	public static void generarWordDesdePlantilla(String plantillaPath, String salidaPath, List<String> pasos, List<String> resultados, String nombreCaso, String prerequisitos, String fecha, String iniciativa, String proyecto) throws IOException {
		try (InputStream is = new FileInputStream(plantillaPath);
			 XWPFDocument doc = new XWPFDocument(is)) {
			for (XWPFParagraph p : doc.getParagraphs()) {
				reemplazarMarcadoresEnParrafo(p, pasos, resultados, nombreCaso, prerequisitos, fecha, iniciativa, proyecto);
			}
			doc.getTables().forEach(table -> {
				table.getRows().forEach(row -> {
					row.getTableCells().forEach(cell -> {
						for (XWPFParagraph p : cell.getParagraphs()) {
							reemplazarMarcadoresEnParrafo(p, pasos, resultados, nombreCaso, prerequisitos, fecha, iniciativa, proyecto);
						}
					});
				});
			});
			eliminarDesdeSiguienteMarcador(doc, pasos, resultados);
			try (FileOutputStream fos = new FileOutputStream(salidaPath)) {
				doc.write(fos);
			}
		}
	}
}

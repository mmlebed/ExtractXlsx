package com.company;

import org.w3c.dom.Document;
import org.w3c.dom.NodeList;
import org.xml.sax.SAXException;

import javax.xml.XMLConstants;
import javax.xml.namespace.NamespaceContext;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.xpath.*;
import java.io.*;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.*;
import java.util.logging.Level;
import java.util.logging.Logger;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;

public class Main {

    public static Map<String, Integer> map;
    public static final int BUFFER_SIZE = 1024;
    public static ArrayList<String> listSheets = new ArrayList<>();
    public static List<String> listStrings = new ArrayList<>();
    public static String ZIP_ARCHIVE;
    public static File dstDir1;
    public static List<String> listR = new ArrayList<>();

    public static void main(String[] args) throws Exception, IOException {
        BufferedReader reader = new BufferedReader(new InputStreamReader(System.in));
        while (true) {
            System.out.print("Введите путь к файлу (например C:\\Папка\\Имя файла.xlsx): ");
            ZIP_ARCHIVE = reader.readLine();
            File file = new File(ZIP_ARCHIVE);
            if (file.isFile()) break;
            else System.out.println("Такого файла нет.");
        }
        Main app = new Main();
        System.out.println(app.print());
        Path p = Paths.get(ZIP_ARCHIVE);
        String file = p.getFileName().toString();
        System.out.println("Input: " + file);
        app.unZip(ZIP_ARCHIVE);
        app.getName();
        DocumentBuilderFactory builderFactory = DocumentBuilderFactory.newInstance();
        builderFactory.setNamespaceAware(true);
        DocumentBuilder builder;
        Document doc = null;
        String string = dstDir1.toString().replace("\\", "/");
        try {
            builder = builderFactory.newDocumentBuilder();
            doc = builder.parse(string + "/xl/sharedStrings.xml");
            XPathFactory xpathFactory = XPathFactory.newInstance();
            XPath xpath = xpathFactory.newXPath();
            xpath = app.prefix(xpath);
            getStrings(doc, xpath);
            for (int i = 0; i < listSheets.size(); i++) {
                String st1 = listSheets.get(i).substring(0, 1).toUpperCase() + listSheets.get(i).substring(1).toLowerCase();
                System.out.println("-Sheet: " + st1.replaceFirst("[.][^.]+$", ""));
                builder = builderFactory.newDocumentBuilder();
                doc = builder.parse(string + "/xl/worksheets/" + listSheets.get(i));
                getR(doc, xpath);
                for (int j = 0; j < listR.size(); j++) {
                    String sss = listR.get(j);
                    map = getR_T(doc, xpath, sss);
                    Iterator<Map.Entry<String, Integer>> iterator = map.entrySet().iterator();
                    while (iterator.hasNext()) {
                        Map.Entry<String, Integer> pair = iterator.next();
                        if (pair.getValue() > 0) {
                            System.out.println("--" + sss + ": " + listStrings.get(Integer.parseInt(pair.getKey())));
                        } else {
                            System.out.println("--" + sss + ": " + Integer.parseInt(pair.getKey()));
                        }
                    }
                }
            }
        } catch (ParserConfigurationException | SAXException | IOException e) {
            e.printStackTrace();
        }
        System.out.println(app.print());
        app.delete(Main.dstDir1);
        System.in.read();
    }

    private void unZip(final String zipFileName) throws Exception {
        byte[] buffer = new byte[BUFFER_SIZE];
        final String dstDirectory = destinationDirectory(zipFileName);
        final File dstDir = new File(dstDirectory);
        this.dstDir1 = dstDir;
        if (!dstDir.exists()) {
            dstDir.mkdir();
        }
        try {
            final ZipInputStream zis = new ZipInputStream(
                    new FileInputStream(zipFileName));
            ZipEntry ze = zis.getNextEntry();
            String nextFileName;
            while (ze != null) {
                nextFileName = ze.getName();
                File nextFile = new File(dstDirectory + File.separator
                        + nextFileName);
                if (ze.isDirectory()) {
                    nextFile.mkdir();
                } else {
                    new File(nextFile.getParent()).mkdirs();
                    try (FileOutputStream fos
                                 = new FileOutputStream(nextFile)) {
                        int length;
                        while ((length = zis.read(buffer)) > 0) {
                            fos.write(buffer, 0, length);
                        }
                    }
                }
                ze = zis.getNextEntry();
            }
            zis.closeEntry();
            zis.close();
        } catch (FileNotFoundException ex) {
            Logger.getLogger(Main.class.getName())
                    .log(Level.SEVERE, null, ex);
        } catch (IOException ex) {
            Logger.getLogger(Main.class.getName())
                    .log(Level.SEVERE, null, ex);
        }
    }

    private String destinationDirectory(final String srcZip) {
        return srcZip.substring(0, srcZip.lastIndexOf("."));
    }

    public void delete(File f) throws IOException {
        if (f.isDirectory()) {
            for (File c : f.listFiles()) {
                delete(c);
            }
        }
        if (!f.delete()) {
            throw new FileNotFoundException("Failed to delete file: " + f);
        }
    }

    public void getName() throws IOException {
        try {
            File folder = new File(dstDir1.toString() + "/xl/worksheets");
            String[] files = folder.list(new FilenameFilter() {

                @Override
                public boolean accept(File folder, String name) {
                    return name.endsWith(".xml");
                }

            });
            for (String fileName : files) {
                listSheets.add(fileName);
            }
        } catch (Exception e) {
        }
    }

    public XPath prefix(XPath xpath) {
        xpath.setNamespaceContext(new NamespaceContext() {
            public String getNamespaceURI(String prefix) {
                if (prefix == null) {
                    throw new NullPointerException("Null prefix");
                } else if ("spreadsheet".equals(prefix)) {
                    return "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
                } else if ("xml".equals(prefix)) {
                    return XMLConstants.XML_NS_URI;
                }
                return XMLConstants.NULL_NS_URI;
            }
            public String getPrefix(String uri) {
                throw new UnsupportedOperationException();
            }
            public Iterator getPrefixes(String uri) {
                throw new UnsupportedOperationException();
            }
        });
        return xpath;
    }

    public static void getR(Document doc, XPath xpath) throws Exception {
        listR.clear();
        try {
            XPathExpression xPathExpression = xpath.compile("//spreadsheet:row/spreadsheet:c/@r");
            NodeList nodes = (NodeList) xPathExpression.evaluate(doc, XPathConstants.NODESET);
            for (int i = 0; i < nodes.getLength(); i++) {
                listR.add(nodes.item(i).getNodeValue());
            }
        } catch (XPathExpressionException e) {
            e.printStackTrace();
        }
    }

    public static void getStrings(Document doc, XPath xpath) throws Exception {
        try {
            XPathExpression xPathExpression = xpath.compile("//spreadsheet:si/spreadsheet:t/text()");
            NodeList nodes = (NodeList) xPathExpression.evaluate(doc, XPathConstants.NODESET);
            for (int i = 0; i < nodes.getLength(); i++) {
                listStrings.add(nodes.item(i).getNodeValue());
            }
        } catch (XPathExpressionException e) {
            e.printStackTrace();
        }
    }

    public static Map<String, Integer> getR_T(Document doc, XPath xpath, String s) throws Exception {
        Map<String, Integer> map = new HashMap<>();
        String listR_T = "";
        String listR_T1 = "";
        try {
            XPathExpression xPathExpression = xpath.compile("//spreadsheet:row/spreadsheet:c[@r = '" + s + "']/spreadsheet:v/text()");
            NodeList nodes = (NodeList) xPathExpression.evaluate(doc, XPathConstants.NODESET);
            if (nodes.getLength() > 0) {
                listR_T = nodes.item(0).getNodeValue();
            }
            XPathExpression xPathExpression1 = xpath.compile("//spreadsheet:row/spreadsheet:c[@r = '" + s + "']/@t");
            NodeList nodes1 = (NodeList) xPathExpression1.evaluate(doc, XPathConstants.NODESET);
            if (nodes1.getLength() > 0) {
                listR_T1 = nodes1.item(0).getNodeValue();
            }
        } catch (XPathExpressionException e) {
            e.printStackTrace();
        }
        if (listR_T1.equals("s")) {
            map.put(listR_T, 1);
        } else {
            map.put(listR_T, 0);
        }
        return map;
    }

    public boolean checkString(String string) {
        try {
            Integer.parseInt(string);
        } catch (Exception e) {
            return false;
        }
        return true;
    }

    public String print() {
        return "---------";
    }

}

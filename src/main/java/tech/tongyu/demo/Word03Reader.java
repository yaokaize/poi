package tech.tongyu.demo;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.Range;

import java.io.*;
import java.util.HashMap;
import java.util.Map;
import java.util.stream.Collectors;

public class Word03Reader {

    public static void main(String[] args) {
        Map<String, Object> map = new HashMap<>();
        map.put("partyName", "重症");
        map = map.entrySet().stream().collect((Collectors.toMap(v -> "${" + v.getKey() + "}", Map.Entry::getValue)));
        try {
            template03("/template.doc", "/confirmation.doc", map);
        } catch (Throwable e) {
            e.printStackTrace();
            while (e.getCause() != null) {
                e = e.getCause();
            }
            throw new RuntimeException(e.getMessage());
        }
    }

    private static void template03(String templatePath, String newPath, Map<String, Object> map) throws IOException {
        InputStream resourceAsStream = Word03Reader.class.getResourceAsStream(templatePath);
        HWPFDocument doc = new HWPFDocument(resourceAsStream);
        replaceDoc2003(doc, map);
        ByteArrayOutputStream ostream = new ByteArrayOutputStream();
        doc.write(ostream);
        File newFile = new File(Word03Reader.class.getResource(newPath).getFile());
        System.out.println(newFile.getAbsoluteFile());
        try (OutputStream outs = new FileOutputStream(newFile)) {
            outs.write(ostream.toByteArray());
        }

    }

    private static void replaceDoc2003(HWPFDocument doc, Map<String, Object> map) {
        Range bodyRange = doc.getRange();
        for (Map.Entry<String, Object> entry : map.entrySet()) {
            bodyRange.replaceText(entry.getKey(), entry.getValue().toString());
        }
    }
}

import de.phip1611.Docx4JSRUtil;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;

import com.spire.doc.*;
import com.spire.doc.documents.TextSelection;
import com.spire.doc.fields.DocPicture;
import com.spire.doc.fields.TextRange;

import java.io.FileNotFoundException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.Map;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.HashMap;


public class DocumentReplacment {

    public static String documntReplace(String nameTemplate, String[][] stringToReplaceAndToReplacment, String outputFileName) throws IOException, Docx4JException {
        HashMap<String, String> replaceMap = new HashMap<>();

        for (String[] i: stringToReplaceAndToReplacment) {
            replaceMap.put("${" + i[0] + "}",i[1]);
        }

        WordprocessingMLPackage template = WordprocessingMLPackage.load(new FileInputStream(nameTemplate));
        Docx4JSRUtil.searchAndReplace(template, replaceMap);
        String basePath = new File("").getAbsolutePath();
        String outputPath = basePath + "\\output";

        if (!Files.exists(Paths.get(outputPath))) {
            new File(outputPath).mkdir();
        }

        outputPath += "\\" + outputFileName;
        File f_output = new File(outputPath);

        if (f_output.exists() && !f_output.isDirectory()) {
            f_output.delete();
        }

        template.save(new File(outputPath));

        return outputPath;
    }

    public static void replaceTextWithImage(String inputPath, String stringToReplace, String imagePath) throws FileNotFoundException, Docx4JException {
        Document document = new Document(inputPath);
        TextSelection[] selections = document.findAllString("${" + stringToReplace + "}", false, true);
        int index = 0;
        TextRange range = null;

        for (Object obj : selections) {
            TextSelection textSelection = (TextSelection) obj;
            DocPicture pic = new DocPicture(document);
            pic.loadImage(imagePath);
            range = textSelection.getAsOneRange();
            index = range.getOwnerParagraph().getChildObjects().indexOf(range);
            range.getOwnerParagraph().getChildObjects().insert(index, pic);
            range.getOwnerParagraph().getChildObjects().remove(range);
        }
        document.saveToFile(inputPath, FileFormat.Docx_2013);

        String del = "Evaluation Warning: The document was created with Spire.Doc for JAVA.";
        WordprocessingMLPackage templateEvaluation = WordprocessingMLPackage.load(new FileInputStream(inputPath));
        Docx4JSRUtil.searchAndReplace(templateEvaluation, Map.of(del, " "));
        templateEvaluation.save(new File(inputPath));
    }
}

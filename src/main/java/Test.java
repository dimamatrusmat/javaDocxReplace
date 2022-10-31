import org.docx4j.openpackaging.exceptions.Docx4JException;

import java.io.IOException;

public class Test {
    public static void main(String[] args) throws IOException, Docx4JException {
        String [] [] stringToReplaceAndToReplacment = new String[][]{
            {"fio", "Матвеев Дмитрий Владимирович"},
            {"special", "Информационные системы"},
            {"fioTo", "Олег Иванович"},
        };

        String output = "Заявление.docx";

        output = DocumentReplacment.documntReplace("template.docx", stringToReplaceAndToReplacment, output);
        DocumentReplacment.replaceTextWithImage(output, "img", "avatar.jpg");
    }
}

package com.devsaki.opensimpleexcel;

import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStreamReader;
import java.io.PrintWriter;
import java.nio.charset.Charset;
import java.util.concurrent.ExecutionException;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.Future;
import java.util.zip.ZipFile;

/**
 * Created by DevSaki on 26/09/2017.
 */
public class Launcher {

    public static final char CHAR_END = (char) -1;

    public static void main(String[] args) throws IOException, ExecutionException, InterruptedException {
        long init = System.currentTimeMillis();
        String excelFile = "C:/Downloads/BigSpreadsheet.xlsx";
        ZipFile zipFile = new ZipFile(excelFile);

        ExecutorService executor = Executors.newFixedThreadPool(4);
        Future<String[]> futureWords = executor.submit(() -> processSharedStrings(zipFile));
        Future<String[][]> futureSheet1 = executor.submit(() -> processSheet1(zipFile));
        String[] words = futureWords.get();
        String[][] sheet1 = futureSheet1.get();

        long end = System.currentTimeMillis();
        System.out.println("Main only open and read: " + (end - init) / 1000);

        ///Doing somethin with the file::Saving as csv
        try(PrintWriter writer = new PrintWriter(excelFile + ".csv", "UTF-8");){
            for (String[] rows : sheet1) {
                for (String cell : rows) {
                    if(cell!=null){
                        writer.append(words[Integer.parseInt(cell)]);
                    }
                    writer.append(";");
                }
                writer.append("\n");
            }
        }
        executor.shutdown();
    }

    public static String[][] processSheet1(ZipFile zipFile) throws IOException {
        String entry = "xl/worksheets/sheet1.xml";
        String[][] result = null;
        char[] dimensionToken = "dimension ref=\"".toCharArray();
        char[] tokenOpenC = "<c r=\"".toCharArray();
        char[] tokenOpenV = "<v>".toCharArray();
        try (BufferedReader br = new BufferedReader(new InputStreamReader(zipFile.getInputStream(zipFile.getEntry(entry)), Charset.forName("UTF-8")))) {
            String dimension = extractNextValue(br, dimensionToken, '"');
            int[] sizes = extractSizeFromDimention(dimension.split(":")[1]);
            br.skip(30); //Between dimension and next tag c exists more or less 30 chars
            result = new String[sizes[0]][sizes[1]];
            String v;
            while((v = extractNextValue(br, tokenOpenC, '"'))!=null){
                int[] indexes = extractSizeFromDimention(v);
                br.skip(7); // t="s">
                String c = extractNextValue(br, tokenOpenV, '<');
                result[indexes[0]-1][indexes[1]-1] = c;
                br.skip(7); ///v></c>
            }
        }
        return result;
    }

    public static int[] extractSizeFromDimention(String dimention){
        StringBuilder sb = new StringBuilder();
        int columns = 0;
        int rows = 0;
        for (char c : dimention.toCharArray()) {
            if(columns==0){
                if(Character.isDigit(c)){
                    columns = convertExcelIndex(sb.toString());
                    sb = new StringBuilder();
                }
            }
            sb.append(c);
        }
        rows = Integer.parseInt(sb.toString());
        return new int[]{rows, columns};
    }

    public static String[] processSharedStrings(ZipFile zipFile) throws IOException {
        String entry = "xl/sharedStrings.xml";
        String[] words = null;
        char[] wordCount = "Count=\"".toCharArray();
        char[] token = "<t>".toCharArray();
        try (BufferedReader br = new BufferedReader(new InputStreamReader(zipFile.getInputStream(zipFile.getEntry(entry)), Charset.forName("UTF-8")))) {
            String uniqueCount = extractNextValue(br, wordCount, '"');
            words = new String[Integer.parseInt(uniqueCount)];
            String nextWord;
            int currentIndex = 0;
            while ((nextWord = extractNextValue(br, token, '<')) != null) {
                words[currentIndex++] = nextWord;
                br.skip(11); //you can skip at least 11 chars "/t></si><si>"
            }
        }
        return words;
    }

    public static int foundNextTokens(BufferedReader br, char[]... tokens) throws IOException {
        char character;
        int[] indexes = new int[tokens.length];
        while ((character = (char) br.read()) != CHAR_END) {
            for (int i = 0; i < indexes.length; i++) {
                if (tokens[i][indexes[i]] == character) {
                    indexes[i]++;
                    if (indexes[i] == tokens[i].length) {
                        return i;
                    }
                } else {
                    indexes[i] = 0;
                }
            }
        }

        return -1;
    }

    public static String extractNextValue(BufferedReader br, char[] token, char until) throws IOException {
        char character;
        StringBuilder sb = new StringBuilder();
        int index = 0;

        while ((character = (char) br.read()) != CHAR_END) {
            if (index == token.length) {
                if (character == until) {
                    return sb.toString();
                } else {
                    sb.append(character);
                }
            } else {
                if (token[index] == character) {
                    index++;
                } else {
                    index = 0;
                }
            }
        }
        return null;
    }

    public static int convertExcelIndex(String index) {
        int result = 0;
        for (char c : index.toCharArray()) {
            result = result * 26 + ((int) c - (int) 'A' + 1);
        }
        return result;
    }

}

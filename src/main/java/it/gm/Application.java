package it.gm;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.charset.Charset;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

/**
 * Created by matthew on 14.07.16.
 */
public class Application {

    static final Logger logger = LoggerFactory.getLogger(Application.class);

    public static void main(String[] args) {
        if (args.length == 2) {
            Path input = Paths.get(args[0]);
            Path output = Paths.get(args[1]);
            if (Files.exists(input)) {
                String fileName = getFileNameWithoutExtension(input);
                if (getFileNameWithoutExtension(output).equals(getFileNameWithoutExtension(input))
                        || output.getFileName().endsWith("csv")
                        ) {
                    // is filePath
                    fileName = getFileNameWithoutExtension(output);
                    output = output.getParent();
                }
                if (!Files.exists(output)) {
                    try {
                        Path created = Files.createDirectories(output);
                    } catch (IOException e) {
                        logger.error(e.getLocalizedMessage());
                        if (logger.isDebugEnabled()) {
                            logger.error(e.getLocalizedMessage(), e);
                        }
                    }
                }
                Spreadsheet s = new Spreadsheet();
                Map<String, String> out = s.convertToCsv(input, '\t');
                // if wb has only one sheet
                if (out.keySet().size() == 1) {
                    String old_key = out.keySet().iterator().next();
                    String value = out.get(old_key);
                    out.remove(old_key);
                    out.put(fileName + ".csv", value);
                } else {
                    //loop through sheets
                    List<String> sheets = new ArrayList<>(out.keySet());
                    for (String name : sheets) {
                        String nextName = fileName + "_" + name + ".csv";
                        out.put(nextName, out.get(name));
                        out.remove(name);
                    }
                }
                writeFileMap(out, output);
            } else {
                logger.info("File: " + args[0] + " does not exist");
                System.exit(1);
            }
        }
    }

    protected static void writeFileMap(Map<String, String> map, Path dir) {
        List<String> sheets = new ArrayList<>(map.keySet());
        for (String name : sheets) {
            Path f = Paths.get(dir.toString(), name);
            logger.info("Writing file " + name + " to directory: " + dir.toAbsolutePath().toString());
            try (FileOutputStream fos = new FileOutputStream(f.toFile())) {
                fos.write(
                        map.get(name).getBytes(
                                Charset.forName("UTF-8")
                        )
                );
                fos.flush();
            } catch (FileNotFoundException e) {
                logger.error(e.getLocalizedMessage(), e);
            } catch (IOException e) {
                logger.error(e.getLocalizedMessage(), e);
            }

        }
    }

    protected static String getFileNameWithoutExtension(Path path) {
        String fileName = path.getFileName().toString();
        int pos = fileName.lastIndexOf(".");
        if (pos > 0) {
            fileName = fileName.substring(0, pos);
        }
        return fileName;
    }
}

package net.cua.export.entity.e7;

import net.cua.export.entity.WaterMark;
import net.cua.export.manager.Const;
import net.cua.export.manager.RelManager;
import net.cua.export.util.ExtBufferedWriter;
import net.cua.export.util.StringUtil;
import org.apache.log4j.Logger;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.sql.ResultSet;
import java.sql.SQLException;

/**
 * Created by wanggq on 2017/9/27.
 */
public class ResultSetSheet extends Sheet {
    private Logger logger = Logger.getLogger(this.getClass().getName());
    private ResultSet rs;
    private boolean copySheet;

    public ResultSetSheet(Workbook workbook) {
        super(workbook);
    }

    public ResultSetSheet(Workbook workbook, String name, HeadColumn[] headColumns) {
        super(workbook, name, headColumns);
    }

    public ResultSetSheet(Workbook workbook, String name, WaterMark waterMark, HeadColumn[] headColumns) {
        super(workbook, name, waterMark, headColumns);
    }

    public ResultSetSheet(Workbook workbook, String name, WaterMark waterMark, HeadColumn[] headColumns, ResultSet rs, RelManager relManager) {
        super(workbook, name, waterMark, headColumns);
        this.rs = rs;
        this.relManager = relManager.clone();
    }

    public void setRs(ResultSet rs) {
        this.rs = rs;
    }

    public ResultSetSheet setCopySheet(boolean copySheet) {
        this.copySheet = copySheet;
        return this;
    }

    @Override
    public void close() {
//        super.close();
        if (rs != null) {
            try {
                rs.close();
            } catch (SQLException e) {
                logger.error(e.getErrorCode(), e);
            }
        }
    }

    @Override
    public void writeTo(Path xl) throws IOException {
        Path worksheets = Paths.get(xl.toString(), "worksheets");
        if (!Files.exists(worksheets)) {
            Files.createDirectory(worksheets);
        }
        String name = getFileName();
        logger.info(getName() + " | " + name);

        for (int i = 0; i < headColumns.length; i++) {
            if (StringUtil.isEmpty(headColumns[i].getName())) {
                headColumns[i].setName(String.valueOf(i));
            }
        }

        boolean paging = false;
        File sheetFile = Paths.get(worksheets.toString(), name).toFile();
        // write date
        try (ExtBufferedWriter bw = new ExtBufferedWriter(new OutputStreamWriter(new FileOutputStream(sheetFile), StandardCharsets.UTF_8))) {
            // Write header
            writeBefore(bw);
            int limit = Const.Limit.MAX_ROWS_ON_SHEET - rows; // exclude header rows
            // Main data
            if (rs != null && rs.next()) {

                // Write sheet data
                if (getAutoSize() == 1) {
                    do {
                        // row >= max rows
                        if (rows >= limit) {
                            paging = !paging;
                            writeRowAutoSize(rs, bw);
                            break;
                        }
                        writeRowAutoSize(rs, bw);
                    } while (rs.next());
                } else {
                    do {
                        // Paging
                        if (rows >= limit) {
                            paging = !paging;
                            writeRow(rs, bw);
                            break;
                        }
                        writeRow(rs, bw);
                    } while (rs.next());
                }
            }

            // Write foot
            writeAfter(bw);

        } catch (IOException e) {
            e.printStackTrace();
        } catch (SQLException e) {
            e.printStackTrace();
        } finally {
            if (rows < Const.Limit.MAX_ROWS_ON_SHEET) {
                close();
            }
        }

        // Delete empty copy sheet
        if (copySheet && rows == 1) {
            logger.info("Delete empty copy sheet");
            workbook.remove(id - 1);
            sheetFile.delete();
            return;
        }

        // resize columns
        boolean resize = false;
        for  (HeadColumn hc : headColumns) {
            if (hc.getWidth() > 0.000001) {
                resize = true;
                break;
            }
        }
        if (getAutoSize() == 1 || resize) {
            autoColumnSize(sheetFile);
        }

        // relationship
        relManager.write(worksheets, name);

        if (paging) {
            int sub;
            if (!copySheet) {
                sub = 0;
            } else {
                sub = Integer.parseInt(this.name.substring(this.name.lastIndexOf('(') + 1, this.name.lastIndexOf(')')));
            }
            String sheetName = this.name;
            if (copySheet) {
                sheetName = sheetName.substring(0, sheetName.lastIndexOf('(') - 1);
            }

            ResultSetSheet rss = clone();
            rss.name = sheetName + " (" + (sub + 1) + ")";
            workbook.insertSheet(id, rss);
            rss.writeTo(xl);
        }

    }

    public ResultSetSheet clone() {
        return new ResultSetSheet(workbook, name, waterMark, headColumns, rs, relManager).setCopySheet(true);
    }
}

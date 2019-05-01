/*
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

package cn.ttzero.excel.entity;

import cn.ttzero.excel.Print;
import cn.ttzero.excel.entity.style.Fill;
import cn.ttzero.excel.entity.style.PatternType;
import cn.ttzero.excel.entity.style.Styles;
import org.junit.Test;

import java.awt.*;
import java.io.IOException;
import java.sql.Connection;
import java.sql.SQLException;

/**
 * Create by guanquan.wang at 2019-04-28 22:47
 */
public class StatementSheetTest extends SQLWorkbookTest {
    @Test public void testWrite() {
        testWrite(false);
    }

    @Test public void testStyleProcessor() {
        testStyleProcessor(false);
    }

    @Test public void testIntConversion() {
        testIntConversion(false);
    }

    // ---- AUTO SIZE

    @Test public void testWriteAutoSize() {
        testWrite(true);
    }

    @Test public void testStyleProcessorAutoSize() {
        testStyleProcessor(true);
    }

    @Test public void testIntConversionAutoSize() {
        testIntConversion(true);
    }

    private void testWrite(boolean autoSize) {
        try (Connection con = getConnection()) {
            new Workbook("statement", author)
                .watch(Print::println)
                .setAutoSize(autoSize)
                .setConnection(con)
                .addSheet("select id, name, age from student order by age"
                    , new Sheet.Column("学号", int.class)
                    , new Sheet.Column("性名", String.class)
                    , new Sheet.Column("年龄", int.class)
                )
                .writeTo(defaultTestPath);
        } catch (SQLException |IOException e) {
            e.printStackTrace();
        }
    }

    private void testStyleProcessor(boolean autoSize) {
        try (Connection con = getConnection()) {
            new Workbook("statement style processor", author)
                .watch(Print::println)
                .setAutoSize(autoSize)
                .setConnection(con)
                .addSheet("select id, name, age from student"
                    , new Sheet.Column("学号", int.class)
                    , new Sheet.Column("性名", String.class)
                    , new Sheet.Column("年龄", int.class)
                        .setStyleProcessor((o, style, sst) -> {
                            int n = (int) o;
                            if (n < 10) {
                                style = Styles.clearFill(style)
                                    | sst.addFill(new Fill(PatternType.solid, Color.orange));
                            }
                            return style;
                        })
                )
                .writeTo(defaultTestPath);
        } catch (SQLException |IOException e) {
            e.printStackTrace();
        }
    }

    private void testIntConversion(boolean autoSize) {
        try (Connection con = getConnection()) {
            new Workbook("test int conversion statement", author)
                .setConnection(con)
                .setAutoSize(autoSize)
                .watch(Print::println)
                .addSheet("select id, name, age from student"
                    , new Sheet.Column("学号", int.class)
                    , new Sheet.Column("姓名", String.class)
                    , new Sheet.Column("年龄", int.class, n -> n > 14 ? "高龄" : n)
                        .setStyleProcessor((o, style, sst) -> {
                            int n = (int) o;
                            if (n > 14) {
                                style = Styles.clearFill(style)
                                    | sst.addFill(new Fill(PatternType.solid, Color.orange));
                            }
                            return style;
                        })
                )
                .writeTo(defaultTestPath);
        } catch (SQLException |IOException e) {
            e.printStackTrace();
        }
    }
}
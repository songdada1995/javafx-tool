package org.example.javafxtool.file;

import net.sf.sevenzipjbinding.IInArchive;
import net.sf.sevenzipjbinding.SevenZip;
import net.sf.sevenzipjbinding.impl.RandomAccessFileInStream;

import java.io.File;
import java.io.IOException;
import java.io.RandomAccessFile;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.Comparator;
import java.util.List;

/**
 * @author songbo
 * @version 1.0
 * @date 2022/8/29 16:30
 */
public class FileUtil {

    /**
     * 删除整个文件夹文件
     *
     * @param path
     * @throws IOException
     */
    public static void deleteWholeFile(String path) throws IOException {
        Files.walk(Paths.get(path))
                .sorted(Comparator.reverseOrder())
                .forEach(v -> {
                    try {
                        Files.delete(v);
                    } catch (IOException e) {
                        e.printStackTrace();
                    }
                });
    }

    /**
     * 查找文件夹内所有文件
     *
     * @param rootFile
     * @param fileList
     * @throws IOException
     */
    public static void listAllFile(File rootFile, List<File> fileList) throws IOException {
        for (File itemFile : rootFile.listFiles()) {
            if (!itemFile.isDirectory()) {
                fileList.add(itemFile);
            } else {
                listAllFile(itemFile, fileList);
            }
        }
    }

    /**
     * 解压文件至目标文件夹
     * 支持zip/rar格式
     *
     * @param sourceFile
     * @param targetFilePath
     * @throws Exception
     */
    public static void uncompressAllFile(File sourceFile, String targetFilePath) throws Exception {
        // 第一个参数是需要解压的压缩包路径，第二个参数参考JdkAPI文档的RandomAccessFile
        RandomAccessFile randomAccessFile = new RandomAccessFile(sourceFile.getPath(), "r");
        IInArchive inArchive = SevenZip.openInArchive(null, new RandomAccessFileInStream(randomAccessFile));
        int[] in = new int[inArchive.getNumberOfItems()];
        for (int i = 0; i < in.length; i++) {
            in[i] = i;
        }

        inArchive.extract(in, false, new UncompressExtractCallback(inArchive, targetFilePath));
        inArchive.close();
        randomAccessFile.close();
    }

    public static void main(String[] args) {
        try {
            uncompressAllFile(new File("F:\\test\\Job\\20240514\\20240514信盒VC.zip"), "F:\\test\\Job\\20240514\\20240514信盒VC_target");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}

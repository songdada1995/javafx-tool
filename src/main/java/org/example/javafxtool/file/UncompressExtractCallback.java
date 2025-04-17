package org.example.javafxtool.file;

import net.sf.sevenzipjbinding.*;

import java.io.*;

/**
 * @author songbo
 * @version 1.0
 * @date 2022/8/29 15:26
 */
public class UncompressExtractCallback implements IArchiveExtractCallback {

    private IInArchive inArchive;
    private String targetDir;

    public UncompressExtractCallback(IInArchive inArchive, String targetDir) {
        this.inArchive = inArchive;
        this.targetDir = targetDir;
    }

    @Override
    public void setCompleted(long arg0) throws SevenZipException {
    }

    @Override
    public void setTotal(long arg0) throws SevenZipException {
    }

    @Override
    public ISequentialOutStream getStream(int index, ExtractAskMode extractAskMode) throws SevenZipException {
        final String path = (String) inArchive.getProperty(index, PropID.PATH);
        final boolean isFolder = (boolean) inArchive.getProperty(index, PropID.IS_FOLDER);

        File targetItemFile = new File(targetDir + File.separator + path);
        if (!targetItemFile.exists() && isFolder) {
            targetItemFile.mkdirs();
        }

        return data -> {
            try {
                if (!isFolder) {
                    writeToFile(targetItemFile, data);
                }
            } catch (Exception e) {
                e.printStackTrace();
            }
            return data.length;
        };
    }

    @Override
    public void prepareOperation(ExtractAskMode arg0) throws SevenZipException {
    }

    @Override
    public void setOperationResult(ExtractOperationResult extractOperationResult) throws SevenZipException {
    }

    public static boolean writeToFile(File file, byte[] msg) throws FileNotFoundException {
        OutputStream fos = null;
        try {
            File parent = file.getParentFile();
            if ((!parent.exists()) && (!parent.mkdirs())) {
                return false;
            }
            fos = new FileOutputStream(file, true);
            fos.write(msg);
            fos.flush();
            return true;
        } catch (FileNotFoundException e) {
            e.printStackTrace();
            return false;
        } catch (IOException e) {
            e.printStackTrace();
            return false;
        } finally {
            try {
                if (null != fos) {
                    fos.close();
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

}
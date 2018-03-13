package com.freshlancers.excelexport;

import android.Manifest;
import android.content.Context;
import android.content.Intent;
import android.content.pm.PackageManager;
import android.net.Uri;
import android.os.Bundle;
import android.support.v4.app.ActivityCompat;
import android.support.v4.content.ContextCompat;
import android.support.v7.app.AppCompatActivity;
import android.util.Log;
import android.widget.Toast;

import com.snatik.storage.Storage;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Arrays;
import java.util.Date;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.Set;

import butterknife.ButterKnife;
import butterknife.OnClick;

public class MainActivity extends AppCompatActivity {

    private static final String TAG = "MainActivity";

    private Storage storage;
    private String newDir;
    private int WRITE_EXTERNAL_STORAGE = 111;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);
        ButterKnife.bind(this);
        //init
        storage = new Storage(getApplicationContext());
        // get external storage
        String path = storage.getExternalStorageDirectory();

        // new dir
        newDir = path + File.separator + "Convert to Excel";
        storage.createDirectory(newDir);

        boolean hasPermission = (ContextCompat.checkSelfPermission(getBaseContext(),
                Manifest.permission.WRITE_EXTERNAL_STORAGE) == PackageManager.PERMISSION_GRANTED);

        if (!hasPermission) {
            ActivityCompat.requestPermissions(MainActivity.this,
                    new String[]{Manifest.permission.READ_EXTERNAL_STORAGE, Manifest.permission.WRITE_EXTERNAL_STORAGE,
                            Manifest.permission.ACCESS_NETWORK_STATE,
                            Manifest.permission.RECORD_AUDIO, Manifest.permission.MODIFY_AUDIO_SETTINGS,
                            Manifest.permission.INTERNET
                    }, WRITE_EXTERNAL_STORAGE);
        }
    }

    @OnClick(R.id.exportButton)
    public void onViewClicked() {
        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet("Sample sheet");

        Map<String, Object[]> data = new LinkedHashMap<>();
        data.put("1", new Object[]{"Emp No.", "Name", "Salary"});
        data.put("2", new Object[]{1, "John", 1500000d});
        data.put("3", new Object[]{2, "Sam", 800000d});
        data.put("4", new Object[]{3, "Dean", 700000d});

        Set<String> keyset = data.keySet();
        int rowNum = 0;
        for (String key : keyset) {
            Row row = sheet.createRow(rowNum++);
            Object[] objArr = data.get(key);
            Log.e(TAG, "onViewClicked: " + key);
            Log.e(TAG, "onViewClicked: " + Arrays.toString(objArr));
            int cellNum = 0;
            for (Object obj : objArr) {
                Cell cell = row.createCell(cellNum++);
                cell.setCellValue(String.valueOf(obj));
                if (obj instanceof Date)
                    cell.setCellValue((Date) obj);
                else if (obj instanceof Boolean)
                    cell.setCellValue((Boolean) obj);
                else if (obj instanceof String)
                    cell.setCellValue((String) obj);
                else if (obj instanceof Double)
                    cell.setCellValue((Double) obj);
            }
        }

        try {
            File xlFile = new File(newDir + File.separator + "new.xls");
            FileOutputStream out = new FileOutputStream(xlFile);
            workbook.write(out);
            out.close();
            Log.e(TAG, "onViewClicked: " + "Excel written successfully..");
            Toast.makeText(this, "Excel saved to " + newDir + File.separator + "new.xls", Toast.LENGTH_LONG).show();
//
//            Intent shareIntent = new Intent(Intent.ACTION_SEND);
//            shareIntent.setType("application/vnd.ms-excel");
//
//            File xlsFile = new File(getFilesDir(), newDir + File.separator + "new.xls");
//            shareIntent.putExtra(Intent.EXTRA_STREAM, Uri.fromFile(xlsFile));
//            startActivity(Intent.createChooser(shareIntent, "Export as "));

           // openFile(MainActivity.this,xlFile);

        } catch (IOException e) {
            e.printStackTrace();
            Log.e(TAG, "onViewClicked: " + e.getMessage());
            Toast.makeText(this, "Excel couldn't be saved to " + newDir + File.separator + "new.xls", Toast.LENGTH_LONG).show();

        }
    }

    public void openFile(Context context, File url) throws IOException {
        // Create URI
        File file = url;
        Uri uri = Uri.fromFile(file);

        Intent intent = new Intent(Intent.ACTION_VIEW);
        // Check what kind of file you are trying to open, by comparing the url with extensions.
        // When the if condition is matched, plugin sets the correct intent (mime) type,
        // so Android knew what application to use to open the file
        if (url.toString().contains(".doc") || url.toString().contains(".docx")) {
            // Word document
            intent.setDataAndType(uri, "application/msword");
        } else if (url.toString().contains(".pdf")) {
            // PDF file
            intent.setDataAndType(uri, "application/pdf");
        } else if (url.toString().contains(".ppt") || url.toString().contains(".pptx")) {
            // Powerpoint file
            intent.setDataAndType(uri, "application/vnd.ms-powerpoint");
        } else if (url.toString().contains(".xls") || url.toString().contains(".xlsx")) {
            // Excel file
            intent.setDataAndType(uri, "application/vnd.ms-excel");
        } else if (url.toString().contains(".zip") || url.toString().contains(".rar")) {
            // WAV audio file
            intent.setDataAndType(uri, "application/x-wav");
        } else if (url.toString().contains(".rtf")) {
            // RTF file
            intent.setDataAndType(uri, "application/rtf");
        } else if (url.toString().contains(".wav") || url.toString().contains(".mp3")) {
            // WAV audio file
            intent.setDataAndType(uri, "audio/x-wav");
        } else if (url.toString().contains(".gif")) {
            // GIF file
            intent.setDataAndType(uri, "image/gif");
        } else if (url.toString().contains(".jpg") || url.toString().contains(".jpeg") || url.toString().contains(".png")) {
            // JPG file
            intent.setDataAndType(uri, "image/jpeg");
        } else if (url.toString().contains(".txt")) {
            // Text file
            intent.setDataAndType(uri, "text/plain");
        } else if (url.toString().contains(".3gp") || url.toString().contains(".mpg") || url.toString().contains(".mpeg") || url.toString().contains(".mpe") || url.toString().contains(".mp4") || url.toString().contains(".avi")) {
            // Video files
            intent.setDataAndType(uri, "video/*");
        } else {
            //if you want you can also define the intent type for any other file

            //additionally use else clause below, to manage other unknown extensions
            //in this case, Android will show all applications installed on the device
            //so you can choose which application to use
            intent.setDataAndType(uri, "*/*");
        }

        intent.addFlags(Intent.FLAG_ACTIVITY_NEW_TASK);
        context.startActivity(intent);
    }
}

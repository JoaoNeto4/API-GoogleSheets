import com.google.api.client.auth.oauth2.Credential;
import com.google.api.client.extensions.java6.auth.oauth2.AuthorizationCodeInstalledApp;
import com.google.api.client.extensions.jetty.auth.oauth2.LocalServerReceiver;
import com.google.api.client.googleapis.auth.oauth2.GoogleAuthorizationCodeFlow;
import com.google.api.client.googleapis.auth.oauth2.GoogleClientSecrets;
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.http.javanet.NetHttpTransport;
import com.google.api.client.json.JsonFactory;
import com.google.api.client.json.gson.GsonFactory;
import com.google.api.client.util.store.FileDataStoreFactory;
import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.SheetsScopes;
import com.google.api.services.sheets.v4.model.AppendValuesResponse;
import com.google.api.services.sheets.v4.model.BatchUpdateValuesRequest;
import com.google.api.services.sheets.v4.model.BatchUpdateValuesResponse;
import com.google.api.services.sheets.v4.model.UpdateValuesResponse;
import com.google.api.services.sheets.v4.model.ValueRange;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.security.GeneralSecurityException;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.List;

public class SheetsQuickstart {
    //private static Sheets sheetsService;
    private static final String APPLICATION_NAME = "Google Sheets API Java Quickstart";
    private static final JsonFactory JSON_FACTORY = GsonFactory.getDefaultInstance();
    private static final String LINK_PLANILHA = "tokens";

    //private static final List<String> SCOPES = Collections.singletonList(SheetsScopes.SPREADSHEETS_READONLY);
    private static final List<String> SCOPES = Collections.singletonList(SheetsScopes.SPREADSHEETS);
    private static final String CREDENTIALS_FILE_PATH = "/credentials.json";///credentials.json dentro de resources
    private Connection conexao;
    private String banco = "jdbc:sqlserver://localhost;databaseName=dbname;useSSL=false;"; 
    /**
     * Creates an authorized Credential object.
     * @param HTTP_TRANSPORT The network HTTP Transport.
     * @return An authorized Credential object.
     * @throws IOException If the credentials.json file cannot be found.
     */

     
    private static Credential getCredentials(final NetHttpTransport HTTP_TRANSPORT) throws IOException {
        // Leitura da credentials.
        InputStream in = SheetsQuickstart.class.getResourceAsStream(CREDENTIALS_FILE_PATH);
        if (in == null) {
            throw new FileNotFoundException("Erro com as credenciais: " + CREDENTIALS_FILE_PATH);
        }
        GoogleClientSecrets clientSecrets = GoogleClientSecrets.load(JSON_FACTORY, new InputStreamReader(in));

        // Requisição de autorização do usuário
        GoogleAuthorizationCodeFlow flow = new GoogleAuthorizationCodeFlow.Builder(
                HTTP_TRANSPORT, JSON_FACTORY, clientSecrets, SCOPES)
                .setDataStoreFactory(new FileDataStoreFactory(new java.io.File(LINK_PLANILHA)))
                .setAccessType("offline")
                .build();
                //System.out.println("teste----> "+ flow.toString());
        LocalServerReceiver receiver = new LocalServerReceiver.Builder().setPort(8888).build();
        return new AuthorizationCodeInstalledApp(flow, receiver).authorize("notJoaos");
    }

    /**** mostra as informações no console *****/
    public void getDataSheets()throws IOException, GeneralSecurityException{
        final NetHttpTransport HTTP_TRANSPORT = GoogleNetHttpTransport.newTrustedTransport();
        final String spreadsheetId = "***************************";//ID contido no link da planilha
        final String range = "A2:E";//Caso tenha Abas: aba1!A2:E
        Sheets service = new Sheets.Builder(HTTP_TRANSPORT, JSON_FACTORY, getCredentials(HTTP_TRANSPORT))
                .setApplicationName(APPLICATION_NAME)
                .build();
        ValueRange response = service.spreadsheets().values()
                .get(spreadsheetId, range)
                .execute();
        List<List<Object>> values = response.getValues();
        if (values == null || values.isEmpty()) {
            System.out.println("A planilha não contém dados!");
        } else {
            System.out.println("Coluna_1, Coluna_2, Coluna_3, Coluna_4, Coluna_5");
            for (List row : values) {
                // Printa as coluna no console
                System.out.printf("%s, %s, %s, %s, %s\n", row.get(0), row.get(1), row.get(2), row.get(3), row.get(4));
            }
        }
    }

    public void insertDataSheets()throws IOException, GeneralSecurityException{  
        // Insere um novo dado na planilh
        final NetHttpTransport HTTP_TRANSPORT = GoogleNetHttpTransport.newTrustedTransport();
        final String spreadsheetId = "*******************************";//
        final String range = "A2:E"; 
        Sheets service = new Sheets.Builder(HTTP_TRANSPORT, JSON_FACTORY, getCredentials(HTTP_TRANSPORT))
                .setApplicationName(APPLICATION_NAME)
                .build();
        ValueRange response = service.spreadsheets().values().get(spreadsheetId, range).execute();
        List<List<Object>> values = Arrays.asList(
                Arrays.asList("eitaaa","6897877","teste1","teste2", "teste3"));
                ValueRange body = new ValueRange()
                .setValues(values);
        AppendValuesResponse result =
                    service.spreadsheets().values().append(spreadsheetId, range, body)
                            .setValueInputOption("USER_ENTERED").setInsertDataOption("INSERT_ROWS").setIncludeValuesInResponse(true)
                            .execute();
            System.out.printf("%d cells appended.", result.getUpdates().getUpdatedCells());
        List<List<Object>> value = response.getValues();
        if (values == null || values.isEmpty()) {
            System.out.println("Sem dados.");
        } else {
            System.out.println("col_1, col_2, col_3, col_4, col_5");
            for (List row : value) {
                System.out.printf("%s, %s, %s, %s, %s\n", row.get(0), row.get(1),row.get(2),row.get(3),row.get(4));
            }
        }
    }

    public void updateDataSheets()throws IOException, GeneralSecurityException{
        // Atualiza dados na planilha Sheets
        final NetHttpTransport HTTP_TRANSPORT = GoogleNetHttpTransport.newTrustedTransport();
        final String spreadsheetId = "**************************";//
        final String range = "A2:E"; //I am intentionally not showing the name and also the link
        Sheets service = new Sheets.Builder(HTTP_TRANSPORT, JSON_FACTORY, getCredentials(HTTP_TRANSPORT))
                .setApplicationName(APPLICATION_NAME)
                .build();
        ValueRange response = service.spreadsheets().values().get(spreadsheetId, range).execute();
        List<List<Object>> values = Arrays.asList(
                Arrays.asList("alteado","6897877","teste1","teste2", "teste3"));
                ValueRange body = new ValueRange()
                .setValues(values);
        UpdateValuesResponse result =
                    service.spreadsheets().values().update(spreadsheetId, "A9:E", body)//usar um range de intervalo ao invez de "String"
                            .setValueInputOption("RAW")
                            .execute();
            System.out.printf("Atualizado");
       
    }


    //*****************************************************************************************************************************************/
    //https://www.tabnine.com/code/java/classes/com.google.api.services.sheets.v4.model.ValueRange
    public void insertDataSheetsLote(String valueInputOption, List<List<Object>> _values) throws IOException,GeneralSecurityException {
        final NetHttpTransport HTTP_TRANSPORT = GoogleNetHttpTransport.newTrustedTransport();
        final String spreadsheetId = "**************************";//
        final String range = "A2:E"; //I am intentionally not showing the name and also the link
        Sheets service = new Sheets.Builder(HTTP_TRANSPORT, JSON_FACTORY, getCredentials(HTTP_TRANSPORT))
        .setApplicationName(APPLICATION_NAME)
        .build();
        // [START sheets_batch_update_values]
        List<List<Object>> values = Arrays.asList(
            Arrays.asList(
                // Cell values ...
            )
             // Additional rows ...
        );
        // [START_EXCLUDE silent]
        values = _values;
        // [END_EXCLUDE]
        List<ValueRange> data = new ArrayList<ValueRange>();
        data.add(new ValueRange()
            .setRange(range)
            .setValues(values));
        // Additional ranges to update ...
        BatchUpdateValuesRequest body = new BatchUpdateValuesRequest()
            .setValueInputOption(valueInputOption)
            .setData(data);
        BatchUpdateValuesResponse result =
            service.spreadsheets().values().batchUpdate(spreadsheetId, body).execute();
        System.out.printf("%d células atualizadas.", result.getTotalUpdatedCells());
        // [END sheets_batch_update_values]
        //https://developers.google.com/sheets/api/reference/rest/v4/ValueRenderOption
        /*********************EXEMPLO DE CHAMADA******************/
        /*
        SheetsQuickstart t = new SheetsQuickstart();
        List<List<Object>> list = new ArrayList<>();
        List<Object> a = new ArrayList<>();
        a.add("c");
        a.add("a");
        List<Object> b = new ArrayList<>();
        b.add("h");
        b.add("b");
        list.add(a);
        list.add(b);
        t.insertDataSheetsLote("RAW", list);
        */
    }

    
    public Connection getConexao() throws Exception{
        conexao = java.sql.DriverManager.getConnection(banco, "user", "password");
        return conexao;
    }

    public void MostraDados(){
        String SQL = "SELECT T0.[ItemCode] FROM  [dbo].[OITM] T0";
        //SELECT top 10 T0.[ItemCode] FROM  [dbo].[OITM] T0
        SheetsQuickstart t = new SheetsQuickstart();
        try {
            Connection conn= t.getConexao();
            PreparedStatement stmt = conn.prepareStatement(SQL);
            ResultSet rs = stmt.executeQuery();
            while (rs.next()) {
                System.out.println("item: "+rs.getString(1));
            }
            conn.close();
        } catch (Exception e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
        
    }
   
    /**
     * Prints the names and majors of students in a sample spreadsheet:
     * https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit
     * @throws Exception
     */
    public static void main(String... args) throws Exception {
       
        String SQL = "SELECT top 10 T0.[ItemCode],T0.[ItemName] FROM  [dbo].[OITM] T0";
        SheetsQuickstart t = new SheetsQuickstart();
        try {
            Connection conn= t.getConexao();
            PreparedStatement stmt = conn.prepareStatement(SQL);
            ResultSet rs = stmt.executeQuery();

            List<List<Object>> list = new ArrayList<>();
           
            while (rs.next()) {
                List<Object> lin = new ArrayList<>();
                lin.add(rs.getString(1));
                lin.add(rs.getString(2));
                list.add(lin);
                //System.out.println("item: "+rs.getString(1));
            }
            t.insertDataSheetsLote("RAW", list);
            conn.close();
        } catch (Exception e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
        
        /*
        SheetsQuickstart t = new SheetsQuickstart();
        List<List<Object>> list = new ArrayList<>();
        List<Object> a = new ArrayList<>();
        a.add("aa");
        a.add("bb");
        List<Object> b = new ArrayList<>();
        b.add("cc");
        b.add("dd");
        List<Object> c = new ArrayList<>();
        c.add("ee");
        c.add("ff");
        list.add(a);
        list.add(b);
        list.add(c);
        t.insertDataSheetsLote("RAW", list);
        */
    }
    
}
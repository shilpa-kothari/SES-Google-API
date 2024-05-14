package SESGoogleSheet;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.security.GeneralSecurityException;
import java.util.Arrays;
import java.util.List;

import com.google.api.client.auth.oauth2.Credential;
import com.google.api.client.extensions.java6.auth.oauth2.AuthorizationCodeInstalledApp;
import com.google.api.client.extensions.jetty.auth.oauth2.LocalServerReceiver;
import com.google.api.client.googleapis.auth.oauth2.GoogleAuthorizationCodeFlow;
import com.google.api.client.googleapis.auth.oauth2.GoogleClientSecrets;
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.http.HttpTransport;
import com.google.api.client.json.JsonFactory;
import com.google.api.client.json.gson.GsonFactory;
import com.google.api.client.util.store.FileDataStoreFactory;

import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.SheetsScopes;
import com.google.api.services.sheets.v4.model.ClearValuesRequest;
import com.google.api.services.sheets.v4.model.ValueRange;

public class GoogleSheetAPI {
    private static final String APPLICATION_NAME = "Google Sheets API";
    private static final JsonFactory JSON_FACTORY =  GsonFactory.getDefaultInstance();
    private static final String TOKENS_DIRECTORY_PATH = "tokens";

    private Sheets service;

    public GoogleSheetAPI() throws IOException, GeneralSecurityException {
        // Authorize using OAuth
        Credential credential = authorize();
        service = new Sheets.Builder(GoogleNetHttpTransport.newTrustedTransport(), JSON_FACTORY, credential)
                .setApplicationName(APPLICATION_NAME)
                .build();
    }

    public void copySheetData(String spreadsheetId, String sourceSheetName, String targetSheetName) throws IOException {
        // Get the source sheet data
        ValueRange response = service.spreadsheets().values()
                .get(spreadsheetId, sourceSheetName + "!A1:Z1000") // adjust the range as needed
                .execute();
        List<List<Object>> values = response.getValues();

        // Clear the target sheet
        ClearValuesRequest requestBody = new ClearValuesRequest();
        service.spreadsheets().values()
                .clear(spreadsheetId, targetSheetName + "!A1:AV131", requestBody) // adjust the range as needed
                .execute();

        // Copy the data to the target sheet
        ValueRange body = new ValueRange().setValues(values);
        service.spreadsheets().values()
                .update(spreadsheetId, targetSheetName + "!A1", body)
                .setValueInputOption("USER_ENTERED")
                .execute();
    }

    private Credential authorize() throws IOException, GeneralSecurityException {
        // Load client secrets
        InputStream in = GoogleSheetAPI.class.getResourceAsStream("/credentials.json");
        GoogleClientSecrets clientSecrets = GoogleClientSecrets.load(JSON_FACTORY, new InputStreamReader(in));

        // Set up authorization code flow
        GoogleAuthorizationCodeFlow flow = new GoogleAuthorizationCodeFlow.Builder(
                GoogleNetHttpTransport.newTrustedTransport(), JSON_FACTORY, clientSecrets,
                Arrays.asList(SheetsScopes.SPREADSHEETS))
                .setDataStoreFactory(new FileDataStoreFactory(new java.io.File(TOKENS_DIRECTORY_PATH)))
                .setAccessType("offline")
                .build();

        // Authorize
        LocalServerReceiver receiver = new LocalServerReceiver.Builder().setPort(8660).build();
        return new AuthorizationCodeInstalledApp(flow, receiver).authorize("user");
    }
}
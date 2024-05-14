
package SESGoogleSheet;

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
import com.google.api.services.drive.Drive;
import com.google.api.services.drive.DriveScopes;
import com.google.api.services.drive.model.File;
import com.google.api.services.drive.model.FileList;
import com.google.api.services.drive.model.Permission;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.security.GeneralSecurityException;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.Collections;
import java.util.List;

import com.google.api.client.http.FileContent;
import com.google.api.client.http.HttpTransport;

/* class to demonstrate use of Drive files list API */
public class GoogleDriveServicenew {
  /**
   * Application name.
   */
  private static final String APPLICATION_NAME = "Google Drive API Java Quickstart";
  /**
   * Global instance of the JSON factory.
   */
  private static final JsonFactory JSON_FACTORY = GsonFactory.getDefaultInstance();
  /**
   * Directory to store authorization tokens for this application.
   */
  private static final String TOKENS_DIRECTORY_PATH = "tokens";

  /**
   * Global instance of the scopes required by this quickstart.
   * If modifying these scopes, delete your previously saved tokens/ folder.
   */
  private static final List<String> SCOPES =
      Collections.singletonList(DriveScopes.DRIVE_FILE);
  private static final String CREDENTIALS_FILE_PATH = "/credentials.json";

  /**
   * Creates an authorized Credential object.
   *
   * @param HTTP_TRANSPORT The network HTTP Transport.
   * @return An authorized Credential object.
   * @throws IOException If the credentials.json file cannot be found.
   */
  private  Credential getCredentials(final NetHttpTransport HTTP_TRANSPORT)
      throws IOException {
    // Load client secrets.
    InputStream in = GoogleDriveServicenew.class.getResourceAsStream(CREDENTIALS_FILE_PATH);
    if (in == null) {
      throw new FileNotFoundException("Resource not found: " + CREDENTIALS_FILE_PATH);
    }
    GoogleClientSecrets clientSecrets =
        GoogleClientSecrets.load(JSON_FACTORY, new InputStreamReader(in));

    // Build flow and trigger user authorization request.
    GoogleAuthorizationCodeFlow flow = new GoogleAuthorizationCodeFlow.Builder(
        HTTP_TRANSPORT, JSON_FACTORY, clientSecrets, SCOPES)
        .setDataStoreFactory(new FileDataStoreFactory(new java.io.File(TOKENS_DIRECTORY_PATH)))
        .setAccessType("offline")
        .build();
    LocalServerReceiver receiver = new LocalServerReceiver.Builder().setPort(8590).build();
    Credential credential = new AuthorizationCodeInstalledApp(flow, receiver).authorize("user");
    //returns an authorized Credential object.
    return credential;
  }

  
  public List<File> getFiles() {
	  
	List<File> files = null;
	try {
		Drive service = getInstance();
		// Print the names and IDs for up to 10 files.
	    FileList result = service.files().list()
	        .setPageSize(10)
	        .setFields("nextPageToken, files(id, name)")
	        .execute();
	     files = result.getFiles();
	    return files;
	} catch (GeneralSecurityException | IOException e) {
		// TODO Auto-generated catch block
		e.printStackTrace();
	}
	return files;
	 
  }
  
  public Drive getInstance() throws GeneralSecurityException, IOException {
      // Build a new authorized API client service.
      final NetHttpTransport HTTP_TRANSPORT = GoogleNetHttpTransport.newTrustedTransport();
      Drive service = new Drive.Builder(HTTP_TRANSPORT, JSON_FACTORY, getCredentials(HTTP_TRANSPORT))
            .setApplicationName(APPLICATION_NAME)
            .build();
      return service;
   }
  
  private String getFolderId(String folderName) {
	  return "1vk5E0TfBiTJVmSp4wUWBxGvQsoVrl_p0";
//	  try {
//	// Search for the folder by name
//      String query = "mimeType='application/vnd.google-apps.folder' and name='" + folderName + "'";
//      String pageToken = null;
//      Drive service = getInstance();
//      do {
//      FileList result = service.files().list().setQ(query).setSpaces("drive").setPageToken(pageToken).execute();
//      List<File> files = result.getFiles();
//      System.out.println(folderName);
//      if (files == null || files.isEmpty()) {
//          System.out.println("No folders found.");
//      } else {
//          for (File folder : files) {
//              System.out.printf("Folder found: %s (%s)\n", folder.getName(), folder.getId());
//              return folder.getId();
//          }
//      }
//      pageToken = result.getNextPageToken();
//      } while (pageToken != null);
//	  } catch (GeneralSecurityException | IOException e) {
//			// TODO Auto-generated catch block
//			e.printStackTrace();
//		}
//		return null;
  }
  
	public String uploadSheetFileToDrive(String srcFilePath) {
		try {
			Drive service = getInstance();
			// Specify the path of the file you want to upload
		    java.io.File filePath = new java.io.File(srcFilePath);

			// Define the file metadata
		    File fileMetadata = new File();
		    fileMetadata.setMimeType("application/vnd.google-apps.spreadsheet");
		    fileMetadata.setParents(Collections.singletonList(getFolderId("backup")));
		    LocalDateTime currentDateTime = LocalDateTime.now();
		    DateTimeFormatter formatter = DateTimeFormatter.ofPattern("yyyy-MM-dd-HH-mm-ss");
			// Format a LocalDate
			String formattedDate = currentDateTime.format(formatter);
			// Get the original file name
	        String originalFileName = filePath.getName();

	        // Get the file name without extension
	        String fileNameWithoutExtension = originalFileName;
	        int lastIndexOfDot = originalFileName.lastIndexOf('.');
	        if (lastIndexOfDot != -1) {
	            fileNameWithoutExtension = originalFileName.substring(0, lastIndexOfDot);
	        }
	 
	        fileMetadata.setName(fileNameWithoutExtension+ "-" + formattedDate);
	        //fileMetadata.setPermissions(Collections.singletonList(createPublicPermission()));

		    // Upload the file
		    FileContent mediaContent = new FileContent("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", filePath);
		    File uploadFile = service.files().create(fileMetadata, mediaContent)
		            .setFields("id")
		            .execute();
		    System.out.println("File ID: " + uploadFile.getId());
		    return uploadFile.getId();
		} catch (GeneralSecurityException | IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return null;
	}
 
	

	public void deleteFile(String fileId) {
		try {
			Drive service = getInstance();
			// Delete the file
	        service.files().delete(fileId).execute();

	        System.out.println("File deleted successfully.");
		} catch (GeneralSecurityException | IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
	}
}
// [END drive_quickstart]
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.security.GeneralSecurityException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.Scanner;

import com.google.api.client.auth.oauth2.Credential;
import com.google.api.client.extensions.java6.auth.oauth2.AuthorizationCodeInstalledApp;
import com.google.api.client.extensions.jetty.auth.oauth2.LocalServerReceiver;
import com.google.api.client.googleapis.auth.oauth2.GoogleAuthorizationCodeFlow;
import com.google.api.client.googleapis.auth.oauth2.GoogleClientSecrets;
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.http.javanet.NetHttpTransport;
import com.google.api.client.json.JsonFactory;
import com.google.api.client.json.jackson2.JacksonFactory;
import com.google.api.client.util.store.FileDataStoreFactory;
import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.SheetsScopes;
import com.google.api.services.sheets.v4.model.UpdateValuesResponse;
import com.google.api.services.sheets.v4.model.ValueRange;

public class Application {

	//var for credentials and connection for google sheet
	private static String APPLICATION_NAME = null;
    private static final JsonFactory JSON_FACTORY = JacksonFactory.getDefaultInstance();
    private static final String TOKENS_DIRECTORY_PATH = "tokens";
    
    private static final List<String> SCOPES = Collections.singletonList(SheetsScopes.SPREADSHEETS);
    private static final String CREDENTIALS_FILE_PATH = "/credentials.json";
	
    //check credentials
    private static Credential getCredentials(final NetHttpTransport HTTP_TRANSPORT) throws IOException {
        // Load client secrets.
        InputStream in = SheetsQuickstart.class.getResourceAsStream(CREDENTIALS_FILE_PATH);
        if (in == null) {
            throw new FileNotFoundException("Resource not found: " + CREDENTIALS_FILE_PATH);
        }
        GoogleClientSecrets clientSecrets = GoogleClientSecrets.load(JSON_FACTORY, new InputStreamReader(in));

        // Build flow and trigger user authorization request.
        GoogleAuthorizationCodeFlow flow = new GoogleAuthorizationCodeFlow.Builder(
                HTTP_TRANSPORT, JSON_FACTORY, clientSecrets, SCOPES)
                .setDataStoreFactory(new FileDataStoreFactory(new java.io.File(TOKENS_DIRECTORY_PATH)))
                .setAccessType("online")
                .build();
        LocalServerReceiver receiver = new LocalServerReceiver.Builder().setPort(8888).build();
        return new AuthorizationCodeInstalledApp(flow, receiver).authorize("user");
    }
    
    //main program
	public static void main(String[] args) throws IOException, GeneralSecurityException {
		Scanner sc = new Scanner(System.in);
    	
		//connection var
        NetHttpTransport HTTP_TRANSPORT = GoogleNetHttpTransport.newTrustedTransport();
        String spreadsheetId = "1qrDnEcWW9lWH5ea__Tn978G0qeBM30onqv7nNi7Hw7Y";
        String range = "D4:F27" ;
        
        //get data for connection and google sheet date for run program
        System.out.println("-------------------------------------Programa para calcular media de 3 notas dos alunos-------------------------------------");
        
        System.out.println("Regras:");
        System.out.println("1 - Media menor que 5, reprovado por nota;");
        System.out.println("2 - Media igual o maior que 5 e menor que 7, Exame final, defindo a nota para aprovacao final com a formula ( 5 <= (m + naf)/2 );");
        System.out.println("3 - Media maior que 7, Aprovado;");
        System.out.println("4 - Numero de faltas maior que 25 por cento de total de aulas, reprovado por falta.");
        
        System.out.println();
        
        System.out.println("Digite o spreadsheetId: ");
        System.out.println("Exemplo: Em um link tipo https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit");
        System.out.println("O spreadsheetId seria: 1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms");
        spreadsheetId = sc.nextLine();
        
        System.out.print("Digite o nome do documento: ");
        APPLICATION_NAME = sc.nextLine();
        
        System.out.print("Digite a linha que comeca os alunos (exemplo:4): ");
        String fRow = sc.nextLine();
        
        System.out.print("Digite a linha que termina os alunos (exemplo:27): ");
        String lRow = sc.nextLine();
        
        System.out.print("Digite a coluna que comeca as notas e incluem as faltas (exemplo:C): ");
        char fColumn = sc.next().charAt(0);
        
        System.out.print("Digite a coluna que termina as notas (exemplo:F): ");
        char lColumn = sc.next().charAt(0);
        
        System.out.print("Digite o numero de aulas (exemplo:60): ");
        int numClass = sc.nextInt();
        
        //calculate the maximum number of absence the class according to the rule
        numClass= numClass * 25 / 100;
        
        //define range for get date from google sheet
        range = fColumn + fRow + ":" + lColumn + lRow;
        
        //get date from google sheet and save to list
        Sheets service = new Sheets.Builder(HTTP_TRANSPORT, JSON_FACTORY, getCredentials(HTTP_TRANSPORT))
                .setApplicationName(APPLICATION_NAME)
                .build();
        ValueRange response = service.spreadsheets().values()
                .get(spreadsheetId, range)
                .execute();
        List<List<Object>> values = response.getValues();
        List<List<Integer>> valuesInt = new ArrayList<List<Integer>>() ;
        if (values == null || values.isEmpty()) {
            System.out.println("No data found.");
        } else {
            for (List row : values) {
            	List<Integer> valuesIntAux = new ArrayList();
            	for(Object x : row) {
            		valuesIntAux.add( Integer.parseInt(x.toString()) );
            	}
            	valuesInt.add(valuesIntAux);
            }
        }
        
        //decrement one to be able to start inserting data on google sheet
        Integer count = Integer.parseInt(fRow) - 1;
        
        //connection to start inserting data on google sheet
        service = new Sheets.Builder(HTTP_TRANSPORT, JSON_FACTORY, getCredentials(HTTP_TRANSPORT))
                .setApplicationName(APPLICATION_NAME)
                .build();
        
        //walks student by student to calculate the average
        for(List row : valuesInt) {
        	int countAvg = 0;
        	Integer classAux = null;
        	
        	//save the absence the class
        	classAux = Integer.parseInt(row.get(0).toString());
        	
        	//calculates the average
        	countAvg = countAvg + Integer.parseInt(row.get(1).toString()) + Integer.parseInt(row.get(2).toString()) + Integer.parseInt(row.get(3).toString());
        	
        	int average = Math.round(countAvg/3);
        	
        	//set "Reprovado por falta" in google sheets if absent much class 
        	if(classAux > numClass) {
        		SpreadsheetSnippets sss = new SpreadsheetSnippets(service);
        		List<List<Object>> situation = new ArrayList();
        		situation.add(new ArrayList<Object>());
        		List<Object> list = new ArrayList();
        		list.add("Reprovado por Falta");
        		situation.add(list);
        		lColumn++;
        		sss.updateValues(spreadsheetId, lColumn+count.toString(), "USER_ENTERED", situation);
        		
        		//set "0" in "Nota para aprovação final" in google sheets if absent much class 
        		sss = new SpreadsheetSnippets(service);
        		List<List<Object>> grade = new ArrayList();
        		grade.add(new ArrayList<Object>());
        		list = new ArrayList();
        		list.add("0");
        		grade.add(list);
        		lColumn++;
        		sss.updateValues(spreadsheetId, lColumn+count.toString(), "RAW", grade);
        		lColumn--;
        		lColumn--;
        	}
        	
        	//set "Aprovado" in google sheets if average bigger 70
        	else if(average >= 70) {
        		
        		SpreadsheetSnippets sss = new SpreadsheetSnippets(service);
        		List<List<Object>> situation = new ArrayList();
        		situation.add(new ArrayList<Object>());
        		List<Object> list = new ArrayList();
        		list.add("Aprovado");
        		situation.add(list);
        		lColumn++;
        		sss.updateValues(spreadsheetId, lColumn+count.toString(), "USER_ENTERED", situation);
        		
        		//set "0" in "Nota para aprovação final" in google sheets if absent much class
        		sss = new SpreadsheetSnippets(service);
        		List<List<Object>> grade = new ArrayList();
        		grade.add(new ArrayList<Object>());
        		list = new ArrayList();
        		list.add("0");
        		grade.add(list);
        		lColumn++;
        		sss.updateValues(spreadsheetId, lColumn+count.toString(), "RAW", grade);
        		lColumn--;
        		lColumn--;
        	}
        	//set "Reprovado" in google sheets if average less 50
        	else if(average < 50) {
        		SpreadsheetSnippets sss = new SpreadsheetSnippets(service);
        		List<List<Object>> situation = new ArrayList();
        		situation.add(new ArrayList<Object>());
        		List<Object> list = new ArrayList();
        		list.add("Reprovado por Nota");
        		situation.add(list);
        		lColumn++;
        		sss.updateValues(spreadsheetId, lColumn+count.toString(), "USER_ENTERED", situation);
        		
        		//set "0" in "Nota para aprovação final" in google sheets if absent much class
        		sss = new SpreadsheetSnippets(service);
        		List<List<Object>> grade = new ArrayList();
        		grade.add(new ArrayList<Object>());
        		list = new ArrayList();
        		list.add("0");
        		grade.add(list);
        		lColumn++;
        		sss.updateValues(spreadsheetId, lColumn+count.toString(), "RAW", grade);
        		lColumn--;
        		lColumn--;
        	}
        	//set "Exame Final" in google sheets if average between 50 and 70
        	else {
        		SpreadsheetSnippets sss = new SpreadsheetSnippets(service);
        		List<List<Object>> situation = new ArrayList();
        		situation.add(new ArrayList<Object>());
        		List<Object> list = new ArrayList();
        		list.add("Exame Final");
        		situation.add(list);
        		lColumn++;
        		sss.updateValues(spreadsheetId, lColumn+count.toString(), "USER_ENTERED", situation);
        		
        		//set grade "Nota para aprovação final" in google sheets according to the formula 5 <= (m + naf)/2 if average between 50 and 70
        		sss = new SpreadsheetSnippets(service);
        		List<List<Object>> grade = new ArrayList();
        		grade.add(new ArrayList<Object>());
        		list = new ArrayList();
        		Integer gradeExam = average - 10;//formula 5 <= (m + naf)/2
        		list.add(gradeExam.toString());
        		grade.add(list);
        		lColumn++;
        		sss.updateValues(spreadsheetId, lColumn+count.toString(), "RAW", grade);
        		lColumn--;
        		lColumn--;
        	}
        	
        	
        	count++;
        }
        
        sc.close();

	}

}

package com.github.meldsza.shortcutCreater;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.io.PrintWriter;
import java.nio.file.Paths;
import java.util.Scanner;

/**
 * Shortcut creator
 */
public class Shortcut {
	/**
	 * Path the shortcut should target
	 */
	public String target;
	/**
	 * Where the shortcut needs to be placed
	 */
	public String shortcutPath;
	/**
	 * WindowStyle (integer value)
	 */
	public int windowStyle;
	/**
	 * Path to icon
	 */
	public String iconLocation;
	/**
	 * Description of the shortcut
	 */
	public String description;
	/**
	 * Working Directory of the target in the shortcut
	 */
	public String workingDirectory;
	/**
	 * Refers to LocalApplicationData folder. Works on Windows 2003 and WinXP+
	 */
	public String localAppdata;
	/**
	 * Refers to ApplicationData folder.
	 */
	public String appdata;
	/**
	 * Refers to %UserProfile%
	 */
	public String userProfile;
	/**
	 * Refers to %UserProfile%\Desktop
	 */
	public String desktop;
	/**
	 * Refers to the Users StartMenu Folder
	 */
	public String startMenu;
	/**
	 * Arguments while invoking the target(if any)
	 */
	public String arguments;
	/**
	 * Set hotkey for shortcut(in String format)
	 * Example : "Ctrl+Alt+F9"
	 */
	public String HotKey;
	/**
	 * Create a instance of the Shortcut class to use it
	 */
	public Shortcut() {
		target = "";
		windowStyle=1;
		iconLocation="";
		description="";
		workingDirectory="";
		arguments="";
		HotKey="";
		getPaths();
	}
	/**
	 * Reads a Shortcut from the file provided.(Uses VBS)
	 * @throws FileNotFoundException If File is not found
	 * @param shortcutfile Provide the existing shortcut file
	 * @throws InterruptedException If the creation was unsuccessfull
	 */
	public Shortcut(File shortcutfile) throws FileNotFoundException, InterruptedException {
		if(! shortcutfile.exists())
		{
			throw new FileNotFoundException();
		}
		else{
		String vbs[] = {
				"set WshShell = WScript.CreateObject(\"WScript.Shell\" )",
				"set oShellLink = WshShell.CreateShortcut(\""+shortcutfile.getAbsolutePath()+"\")",
				"WScript.Echo oShellLink.TargetPath ",
				"WScript.Echo oShellLink.WindowStyle ",
				"WScript.Echo oShellLink.Arguments ",
				"WScript.Echo oShellLink.Hotkey ",
				"WScript.Echo oShellLink.IconLocation ",
				"WScript.Echo oShellLink.Description ",
				"WScript.Echo oShellLink.WorkingDirectory "
			};

		shortcutPath=shortcutfile.getAbsolutePath();
		getShortcut(vbs);
		}
		getPaths();
	}
	public void getPaths() {
		// TODO Auto-generated method stub
		userProfile= exec("UserProfile");
		appdata= exec("AppData");
		localAppdata= exec("LocalAppData");
		if(localAppdata.equalsIgnoreCase("%LocalAppData%"))
		{
			localAppdata = Paths.get(userProfile,"Local Settings\\Application Data").toString();
		}
		if((new File(Paths.get(userProfile,"Start Menu").toString())).exists())
		{
			startMenu = Paths.get(userProfile,"Start Menu").toString();
		}
		else
		{
			startMenu = Paths.get(userProfile,"AppData\\Roaming\\Microsoft\\Windows\\Start Menu").toString();
		}
		desktop =  Paths.get(userProfile,"Desktop").toString();
	}
	private String exec(String cmd)
	{
		Process p=null;
		try {
			p = (new ProcessBuilder("cmd", "/C echo %"+cmd+"%")).start();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		InputStream out = p.getInputStream();
		Scanner s = new Scanner(out).useDelimiter("\\A");
		String st = s.nextLine().replace("\"","");
		s.close();
		try {
			out.close();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return st;
		}
	private void getShortcut(String[] vbs) throws FileNotFoundException, InterruptedException
	{
		PrintWriter fout = new PrintWriter(new File("vbs.vbs"));
		for (String a : vbs) {
			fout.println(a);
		}

		fout.close();
		Process p=null;
		try {
			p = (new ProcessBuilder("cscript","//NoLogo", "vbs.vbs")).start();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		p.waitFor();
		new File("vbs.vbs").delete();
		InputStream out = p.getInputStream();
		System.out.println();
		Scanner s = new Scanner(out).useDelimiter("\\A");
		target = s.nextLine().replace("\"","");
		windowStyle = Integer.parseInt(s.nextLine());
		arguments = s.nextLine().replace("\"","");
		HotKey = s.nextLine().replace("\"","");
		iconLocation = s.nextLine().replace("\"","");
		description = s.nextLine().replace("\"","");
		workingDirectory = s.nextLine().replace("\"","");
		s.close();
		try {
			out.close();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		}

	/**
	 * Creates the Shortcut (saves it on the FileSystem)
	 * It uses VBS to create the shortcut
	 */
	public void createschortcut() throws IOException, InterruptedException
	{
		String vbs[] = {
			"set WshShell = WScript.CreateObject(\"WScript.Shell\" )",
			"Set FSO = CreateObject(\"Scripting.FileSystemObject\")",
			"set oShellLink = WshShell.CreateShortcut(\""+shortcutPath+"\")",
			"oShellLink.TargetPath = \""+target+"\"",
			"oShellLink.WindowStyle = "+windowStyle,
			"oShellLink.IconLocation = \""+iconLocation+"\"",
			"oShellLink.Description = \""+description+"\"",
			"oShellLink.WorkingDirectory = \""+workingDirectory+"\"",
			"oShellLink.Hotkey = \""+HotKey+"\"",
			//"oShellLink.Arguments = \""+arguments+"\"",
			"oShellLink.Save"
		};
		File path = new File(Paths.get(shortcutPath).getParent().toString());
		if(! path.exists())
		{
			try {
				path.createNewFile();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}

			PrintWriter fout = new PrintWriter(new File("vbs.vbs"));
			for (String a : vbs) {
				fout.println(a);
			}

			fout.close();

			Process p = Runtime.getRuntime().exec( "wscript vbs.vbs" );
			p.waitFor();
			(new File("vbs.vbs")).delete();
	}
}
import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;
import org.junit.Before;
import org.junit.Test;

import java.io.IOException;


public class Test_SapExample {
    ActiveXComponent Session;


    @Before
    public  void ConnectionSAP()  {
        // -Variables------------------------------------------------------
        ActiveXComponent SAPROTWr, GUIApp, Connection = null;
        Dispatch ROTEntry;
        Variant ScriptEngine = null;
        String systemName="";


//	    	//Opening the SAP Logon
        String sapLogonPath = "\"C:\\Program\fFiles\f(x86)\\SAP\\FrontEnd\\SAPgui\\saplogon.exe\"";
        try {
            Runtime.getRuntime().exec(sapLogonPath);
        } catch (IOException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }

        for (int i = 0; i < 10; i++) {
            try {
                // -Set SapGuiAuto = GetObject("SAPGUI")---------------------------
                SAPROTWr = new ActiveXComponent("SapROTWr.SapROTWrapper");
                System.out.println("SAPROTWr: " + SAPROTWr);
                ROTEntry = SAPROTWr.invoke("GetROTEntry", "SAPGUI").toDispatch();
                System.out.println("ROTEntry: " + ROTEntry);
                // -Set application = SapGuiAuto.GetScriptingEngine------------
                ScriptEngine = Dispatch.call(ROTEntry, "GetScriptingEngine");
                System.out.println("ScriptEngine: " + ScriptEngine);
                break;
            } catch (Exception e) {
                System.out.println(i);
            }

            try {
                Thread.sleep(500);
            } catch (InterruptedException e) {
                e.printStackTrace();
            }
        }

        GUIApp = new ActiveXComponent(ScriptEngine.toDispatch());
        System.out.println("GUIApp: " + GUIApp);
        // SAP Connection Name
        String sapConnectionName = systemName; // this is the name of the connection in SAP Logon
        try {
            Connection = new ActiveXComponent(GUIApp.invoke("OpenConnection", sapConnectionName).toDispatch());
        } catch (Exception e) {
            System.out.println("Ошибка в названии системы SAP: " + systemName);
        }

        //-Set connection = application.Children(0)-------------------
        Session = new ActiveXComponent(Connection.invoke("Children", 0).toDispatch());
        System.out.println("Нашел окно");
        System.out.println(Session);
    }

    @Test
    public void test1(){
        String id="", text="";
        ActiveXComponent Obj;

        Obj = new ActiveXComponent(Session.invoke("FindById", id).toDispatch());
        Obj.setProperty("Text", text);

        Obj = new ActiveXComponent(Session.invoke("FindById", id).toDispatch());
        Obj.invoke("Press");
    }


}

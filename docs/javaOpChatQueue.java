import java.io.IOException;
import java.util.Properties;
import javax.swing.JComponent;
import javax.swing.JCheckBox;

import callback_interfaces.PluginCallbackRegister;
import callback_interfaces.PublicExposedFunctions;
import callback_interfaces.StaticExposedFunctions;
import exceptions.CommandUsedIllegally;
import exceptions.CommandUsedImproperly;
import exceptions.PluginException;
import plugin_interfaces.CommandCallback;
import plugin_interfaces.GenericPluginInterface;
import plugin_interfaces.OutgoingTextCallback;
import plugin_interfaces.PacketCallback;
import util.BNetPacket;
import util.gui.JTextFieldNumeric;
/*
 * Created on Dec 14, 2004
 * By iago
 */

/**
 * @author iago
 *
 */
public class PluginMain extends GenericPluginInterface implements OutgoingTextCallback, CommandCallback, PacketCallback
{
    private long lastSent = System.currentTimeMillis();

    private PublicExposedFunctions out;
    
    private int credits;

    public void load(StaticExposedFunctions staticFuncs)
    {
    }

    public void activate(PublicExposedFunctions out, PluginCallbackRegister register)
    {
        this.out = out;
        register.registerOutgoingTextPlugin(this, null);
        
        register.registerCommandPlugin(this, "clearqueue", 0, false, "AN", "", "Clears the current queue of outgoing messages, and resets timers.", null);
        register.registerCommandPlugin(this, "testqueue", 0, false, "M", "<size>", "Sends 250 messages of the specified size", null);
        
        register.registerIncomingPacketPlugin(this, SID_FLOODDETECTED, null);
        
//        if (out.getLocalSetting(getName(), "max credits").equalsIgnoreCase("800"))
//        {
//            out.putLocalSetting(getName(), "cost - packet", "190");
//            out.putLocalSetting(getName(), "cost - byte", "12");
//            out.putLocalSetting(getName(), "cost - byte over threshold", "15");
//            out.putLocalSetting(getName(), "starting credits", "750");
//            out.putLocalSetting(getName(), "threshold bytes", "65");
//            out.putLocalSetting(getName(), "max credits", "750");
//            out.putLocalSetting(getName(), "credit rate", "8");
//        }
//        


        credits = Integer.parseInt(out.getLocalSettingDefault(getName(), "starting credits", "750"));
    }

    public void deactivate(PluginCallbackRegister register)
    {
    }

    public String getName()
    {
        return "Antiflood";
    }
    public String getVersion()
    {
        return "1.4";
    }

    public String getAuthorName()
    {
        return "iago";
    }

    public String getAuthorWebsite()
    {
        return "http://www.javaop.com";
    }

    public String getAuthorEmail()
    {
        return "iago@valhallalegends.com";
    }

    public String getShortDescription()
    {
        return "Provides an anti-flood algorithm.";
    }

    public String getLongDescription()
    {
        return "This is an anti-flood algorithm based on one that I wrote, which was based loosely " + 
        	"on Adron's.  I've never flooded off with this algorithm before, but I hear that it gets " + 
        	"unbearably slow after awhile.  It works well for my uses, though, if anybody has a better " + 
        	"one that they feel like porting to Java, be my guest and I'll include it in releases.";
        	
    }

    public Properties getSettingsDescription()
    {
        Properties p = new Properties();
        p.put("debug", "This will show the current delay and the current number of credits each message, in case you want to find-tune it.");
        p.put("prevent flooding", "It's a very bad idea to turn this off -- if you do, it won't try to stop you from flooding.");
        p.put("cost - packet", "This is the amount of credits 'paid' for each sent packet.");
        p.put("cost - byte", "WARNING: I don't recommend changing ANY of the settings for anti-flood.  But if you want to tweak, you can.  This is the number of credits 'paid' for each byte.");
        p.put("cost - byte over threshold", "This is the amount of credits 'paid' for each byte after the threshold is reached.");
        p.put("starting credits", "This is the number of credits you start with.");
        p.put("threshold bytes", "This is the length of a message that triggers the treshold (extra delay).");
        p.put("max credits", "This is the maximum number of credits that the bot can have.");
        p.put("credit rate", "This is the amount of time (in milliseconds) it takes to earn one credit.");
        
        return p;
    }
    
    public Properties getDefaultSettingValues()
    {
        Properties p = new Properties();
        p.put("debug", "false");
        
        p.put("prevent flooding", "true");
        p.put("cost - packet", "250");
        p.put("cost - byte", "15");
        p.put("cost - byte over threshold", "20");
        p.put("starting credits", "200");
        p.put("threshold bytes", "65");
        p.put("max credits", "800");
        p.put("credit rate", "10");
        
        return p;
    }

	public JComponent getComponent(String settingName, String value)
	{
		if(settingName.equalsIgnoreCase("debug") || settingName.equalsIgnoreCase("prevent flooding"))
        {
			return new JCheckBox("", value.equalsIgnoreCase("true"));
        }
        else if(settingName.equalsIgnoreCase("cost - packet") || 
                settingName.equalsIgnoreCase("cost - byte") || 
                settingName.equalsIgnoreCase("cost - byte over threshold") || 
                settingName.equalsIgnoreCase("starting credits") ||
                settingName.equalsIgnoreCase("thresholdBytes") || 
                settingName.equalsIgnoreCase("max credits") ||
                settingName.equalsIgnoreCase("credit rate"))
        {
            return new JTextFieldNumeric(value);
        }
		return null;
	
	}

    
    public Properties getGlobalDefaultSettingValues()
    {
        Properties p = new Properties();
        return p;
    }
    public Properties getGlobalSettingsDescription()
    {
        Properties p = new Properties();
        return p;
    }
    public JComponent getGlobalComponent(String settingName, String value)
    {
        return null;
    }
    
    

    public boolean isRequiredPlugin()
    {
        return true;
    }
    
    public String queuingText(String text, Object data)
    {
        return text;
    }

    public void queuedText(String text, Object data)
    {
    }

    public String nextInLine(String text, Object data)
    {
        return text;
    }
    

    /** Note on implementation: This will assume that all previous packets have already been sent.  Don't call this multiple
     * time in a row and hope to get a good result! */
    public long getDelay(String text, Object data)
    {
        if(out.getLocalSettingDefault(getName(), "prevent flooding", "true").equals("false"))
            return 0;
        
        boolean debug               = out.getLocalSettingDefault(getName(), "debug", "false").equalsIgnoreCase("true");
        int packetCost              = Integer.parseInt(out.getLocalSettingDefault(getName(), "cost - packet", "250"));
        int byteCost                = Integer.parseInt(out.getLocalSettingDefault(getName(), "cost - byte", "15"));
        int byteOverThresholdCost   = Integer.parseInt(out.getLocalSettingDefault(getName(), "cost - byte over threshold", "20"));
        int thresholdBytes          = Integer.parseInt(out.getLocalSettingDefault(getName(), "threshold bytes", "65"));
        int maxCredits              = Integer.parseInt(out.getLocalSettingDefault(getName(), "max credits", "800"));
        int creditRate              = Integer.parseInt(out.getLocalSettingDefault(getName(), "credit rate", "10"));
        
        // Add the credits for the elapsed time
        if(credits < maxCredits)
        {
            credits += (System.currentTimeMillis() - lastSent) / creditRate;
            
            if(credits > maxCredits)
            {
                if(debug)
                    out.systemMessage(DEBUG, "Maximum anti-flood credits reached (" + maxCredits + ")");
                credits = maxCredits;
            }
        }
        
        lastSent = System.currentTimeMillis();

        // Get the packet's "cost"
        int thisByteDelay = byteCost;
        
        if(text.length() > thresholdBytes)
            byteCost = byteOverThresholdCost;

        int thisPacketCost = packetCost + (thisByteDelay * text.length());

        if(debug)
            out.systemMessage(DEBUG, "Cost for this packet = " + thisPacketCost);

        // Check how long this packet will have to wait
        int requiredDelay = 0;
        // If we can't "afford" the packet, figure out how much time we'll have to wait
        if(credits < 0)
            requiredDelay = -credits * creditRate;
        
//        if(thisPacketCost > credits)
//            requiredDelay = -((credits - thisPacketCost) * thisByteDelay);
        
        // Deduct this packet from the credits
        credits -= thisPacketCost;
        
        if(debug)
            out.systemMessage(DEBUG, "Remaining credits: " + credits + "; Delay: " + requiredDelay);
        
        //System.out.println(requiredDelay);

        return requiredDelay;
    }

    public boolean sendingText(String text, Object data)
    {
        return true;
    }

    public void sentText(String text, Object data)
    {
    }

    public void commandExecuted(String user, String command, String[] args, int loudness, Object data) throws PluginException, IOException, CommandUsedIllegally, CommandUsedImproperly
    {
        if(command.equalsIgnoreCase("clearqueue"))
        {
//            firstMessage = 0;
//            sentBytes = 0;
//            sentPackets = 0;
//            totalDelay = 0;
            
            credits = 0;
            lastSent = System.currentTimeMillis();
            
            out.clearQueue();
            out.sendTextUser(user, "Queue cleared", loudness);
        }
        else if(command.equalsIgnoreCase("testqueue"))
        {
            int size = 15;
            if(args.length > 0)
                size = Integer.parseInt(args[0]);
            
            String s = "";
            for(int i = 0; i < size; i++)
                s += 'a';
            for(int i = 0; i < 250; i++)
                out.sendText(s);
        }
        else
        {
            out.sendTextUser(user, "An error occurred in the AntiFlood plugin - unknown command -- please contact iago.", loudness);
        }
    }

    public BNetPacket processingPacket(BNetPacket buf, Object data) throws IOException, PluginException
    {
        return buf;
    }

    public void processedPacket(BNetPacket buf, Object data) throws IOException, PluginException
    {
        if(buf.getCode() == SID_FLOODDETECTED)
        {
            out.systemMessage(ERROR, "You flooded off Battle.net!  If you HAVEN'T tweaked the Anti-flood settings, please send iago the last commands/conversation the bot had!");
        }
    }
}

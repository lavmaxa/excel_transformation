import java.io.*;

public class Main {

    public static void main(String[] args) throws IOException, InterruptedException {
        System.out.println("Framefinder started");
        Senbazuru sbz = new Senbazuru(args[0]);
        System.out.println("Framefinder finished");
        System.out.println("Formatting started");
        Formatting frm = new Formatting(args[1], args[2], args[3]);
        System.out.println("Formatting finished");
        System.out.println("Checking started");
        Checking chck = new Checking(args[1],args[3],args[4], args[2]);
        System.out.println("Checking finished");
    }
}
import org.apache.commons.io.IOUtils;

import java.io.IOException;

public class Senbazuru {
    public Senbazuru(String par) throws IOException, InterruptedException {
        ProcessBuilder process1 = new ProcessBuilder().command("py", par);
        Process p = process1.start();
        process1.redirectErrorStream(true);
        IOUtils.closeQuietly(p.getOutputStream());
        IOUtils.copy(p.getInputStream(), System.out);
        IOUtils.closeQuietly(p.getInputStream());
        int returnVal = p.waitFor();
    }
}

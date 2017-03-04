import org.junit.runner.JUnitCore;
import org.junit.runner.Result;
import org.junit.runner.notification.Failure;

/**
 * Created by zark on 17/2/27.
 */
public class TestRunner {
    public  static void main(String[] args){
        //Result result = JUnitCore.runClasses(TestMessageUtil.class);
        Result result  = JUnitCore.runClasses(TestPoiExcelWriter.class);
        for (Failure failure : result.getFailures()){
            System.out.println(failure.toString());
        }
        System.out.println(result.wasSuccessful());
    }
}

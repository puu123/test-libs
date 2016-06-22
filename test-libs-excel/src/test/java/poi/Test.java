package poi;

import org.databene.contiperf.PerfTest;
import org.databene.contiperf.Required;
import org.databene.contiperf.junit.ContiPerfRule;
import org.junit.Rule;

public class Test {

	@Rule
	public ContiPerfRule i = new ContiPerfRule();

	@org.junit.Test
	@PerfTest(invocations = 1000, threads = 20)
	@Required(max = 1200, average = 250)
	public void test1() throws Exception {
		Thread.sleep(200);
	}
}

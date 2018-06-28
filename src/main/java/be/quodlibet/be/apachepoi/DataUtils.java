package be.quodlibet.be.apachepoi;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.time.ZonedDateTime;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.Date;
import java.util.List;
import java.util.Random;
import java.util.logging.Level;
import java.util.logging.Logger;

/**
 *
 * @author Dries Horions <dries@quodlibet.be>
 */

public class DataUtils
{

    static List<String> cities = Arrays.asList("Brussels", "Antwerp", "Ghent", "Heusden-Zolder", "Lummen", "Schulen", "San Diego");
    static SimpleDateFormat dateFormatter = new SimpleDateFormat("dd/MM/yyyy hh:mm:ss");
    static Random rn = new Random();
    /**
     * Return a list of data records that can be used in our samples
     *
     * @param nr
     * @param durationFormula
     * @param speedFormula
     * @return
     */
    public static List<List<Object>> getRandomRunningResults(int nr, String durationFormula, String speedFormula)
    {
        List<List<Object>> data = new ArrayList();
        for (int i = 0; i < nr; i++) {
            try {
                //create a plausible looking random result for a run
                //Random start date
                Date runStart = getRandomDate("01/01/2018 09:00:00", "01/05/2018 23:59:59");
                //Random distance between 5 and 15 km
                double distance = (5 + rn.nextInt(9)) + rn.nextDouble();
                //Random speed between 9 and 15 km/h
                double speed = 9 + rn.nextInt(5) + rn.nextDouble();
                //Calculate the end time based on speed and distance
                LocalDateTime re = runStart.toInstant().atZone(ZoneId.systemDefault()).toLocalDateTime();
                double seconds = (distance / speed) * 60 * 60;
                re = re.plusSeconds((long) seconds);
                ZonedDateTime zdt = re.atZone(ZoneId.systemDefault());
                Date runEnd = Date.from(zdt.toInstant());
                //A random city
                String city = cities.get(rn.nextInt(cities.size() - 1));
                data.add(Arrays.asList(runStart, runStart, runEnd, runEnd, city, distance, durationFormula, speedFormula));

            }
            catch (ParseException ex) {
                Logger.getLogger(DataUtils.class.getName()).log(Level.SEVERE, null, ex);
            }
        }
        //Sort the list by date by using our own comparator
        Collections.sort(data, (List<Object> lhs, List<Object> rhs) -> {
            Date leftDate = (Date) lhs.get(0);
            Date rightDate = (Date) rhs.get(0);
            if (leftDate.before(rightDate)) {
                return -1;
            }
            if (leftDate.after(rightDate)) {
                return 1;
            }
            return 0;
        });
        return data;
    }

    public static List<List<Object>> getRandomRunningResults(int nr, String durationFormula, String speedFormula, String monthFormula)
    {
        List<List<Object>> data = getRandomRunningResults(nr, durationFormula, speedFormula);
        List<List<Object>> copyList = new ArrayList();
        for (List<Object> record : data) {
            List<Object> copy = new ArrayList(record);
            //Add a field to each record with the Month Formula
            copy.add(monthFormula);
            copyList.add(copy);
        }
        return copyList;
    }

    private static Date getRandomDate(String start, String end) throws ParseException
    {
        Date std = dateFormatter.parse(start);
        Date ste = dateFormatter.parse(end);
        long diff = ste.getTime() - std.getTime() + 1;
        return new Date(std.getTime() + (long) (Math.random() * diff));

    }
}

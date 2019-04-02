import java.util.regex.Pattern;

public class Test {

    public static void main(String[] args) {



        String name = "Sadig    naibbayli";
        String value = "Sadig naibbayli Faig";

      name =   name.replaceAll("\\s+"," ");
        System.out.println(name);

        if (Pattern.matches(name+"\\s*\\w+",value)){
            System.out.println("matches");
        }





    }


}

package org.x.dbstructtoword;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.CommandLineRunner;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

@SpringBootApplication
public class Application implements CommandLineRunner {
    @Autowired
    private PanWeiDbStructToWordTool dbStructToWordTool;

    public static void main(String[] args) {
        SpringApplication.run(Application.class, args);
    }

    @Override
    public void run(String... args) {
        if (args.length < 1) {
            printUsage();
            return;
        }

        String schemaName = args[0];
        if (args.length > 1) {
            String tableName = args[1];
            String tableComment = (args.length > 2) ? args[2] : null;
            dbStructToWordTool.writeWord(schemaName, tableName, tableComment);
        } else {
            dbStructToWordTool.writeWord(schemaName);
        }

    }

    private void printUsage() {
        System.out.println("usage: <command> <schema_name> [table_name] [table_comment]");
    }
}

package com.classes.connection;

import java.io.File;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;

public class conexion {

    public Connection conectarSQL() {
        Connection con = null;
        try {
            con = DriverManager.getConnection("jdbc:sqlite:" + new File("tools/db/database.db"));
        } catch (SQLException e) {
            e.printStackTrace();
        }
        return con;
    }

}


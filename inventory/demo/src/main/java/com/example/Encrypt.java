package com.example;
import org.mindrot.jbcrypt.BCrypt;

public class Encrypt {

    // Pre-hashed passwords (generate once)
    private static final String ADMIN_HASH = BCrypt.hashpw("admin", BCrypt.gensalt());
    private static final String USER_HASH  = BCrypt.hashpw("user", BCrypt.gensalt());

    // Verify password
    public static boolean verifyAdminPassword(String enteredPassword) {
        return BCrypt.checkpw(enteredPassword, ADMIN_HASH);
    }

    public static boolean verifyUserPassword(String enteredPassword) {
        return BCrypt.checkpw(enteredPassword, USER_HASH);
    }
}

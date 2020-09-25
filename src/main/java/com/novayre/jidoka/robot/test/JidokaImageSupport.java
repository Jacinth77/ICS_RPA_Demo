


package com.novayre.jidoka.robot.test;

import java.io.IOException;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Map;

import com.novayre.jidoka.client.api.IJidokaServer;
import com.novayre.jidoka.client.api.IRobot;
import com.novayre.jidoka.client.api.JidokaFactory;
import com.novayre.jidoka.falcon.api.FalconImageOptions;
import com.novayre.jidoka.falcon.api.IFalcon;
import com.novayre.jidoka.falcon.api.IFalconImage;
import com.novayre.jidoka.windows.api.IWindows;

/**
 * Jidoka image support
 *
 * @author Jidoka Team
 *
 */
public class JidokaImageSupport {

    /**
     * Instance
     */
    private static JidokaImageSupport instance;

    /**
     * Server
     */
    private final IJidokaServer<?> server;

    /**
     * Falcon
     */
    private final IFalcon falcon;

    /**
     * Images
     */
    private final Map<String, IFalconImage> images = new HashMap<>();

    /**
     * Constructor
     * @param robot	robot
     */
    public JidokaImageSupport(IRobot robot) {

        server = JidokaFactory.getServer();

        falcon = IFalcon.getInstance(robot, IWindows.getInstance(robot));
    }

    /**
     * Returns instance of class
     * @param robot	robot
     * @return
     */
    public static JidokaImageSupport getInstance(IRobot robot) {

        if (instance == null) {
            instance = new JidokaImageSupport(robot);
        }

        return instance;
    }

    /**
     * Returns a falcon image from specified path
     * @param path
     * @return
     * @throws IOException
     */
    public IFalconImage getImage(Path path) throws IOException {

        IFalconImage image = images.get(path.toString());

        if (image == null) {

            image = falcon.getImage(new FalconImageOptions().description(path.getFileName().toString())
                    .file(path.toFile()).colorTolerance(.05f));

            images.put(path.toString(), image);
        }

        image.setPointsWhereFound(new ArrayList<>());

        return image;
    }

    /**
     * Returns a falcon image referencing image resource test.png
     * @return
     * @throws IOException
     */
    public IFalconImage getTestPng() throws IOException {
        return getImage(Paths.get(server.getCurrentDir(), "test.png"));
    }

    public IFalconImage getDLPng() throws IOException {
        return getImage(Paths.get(server.getCurrentDir(), "DL.png"));
    }

    public IFalconImage getPermitPng() throws IOException {

        return getImage(Paths.get(server.getCurrentDir(), "Permit.png"));
    }

    public IFalconImage getpassportPng() throws IOException {
        return getImage(Paths.get(server.getCurrentDir(), "passport.png"));
    }

}

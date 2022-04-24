import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.io.*;

// Singleton
public class Program {

  private static Program program;
  private ExportLogic exportLogic;
  private MainWindowForm mainWindowForm;
  private JFrame mainFrame;
  private File teilnehmerFile;

  private Program(){

    teilnehmerFile = new File("Teilnehmer.txt");

    mainFrame = new JFrame("Report Generator");
    mainWindowForm = new MainWindowForm();
    exportLogic = new ExportLogic();

    mainFrame.setContentPane(mainWindowForm.getMainPanel());
    mainFrame.setSize(800,600);
    mainFrame.setDefaultCloseOperation(WindowConstants.EXIT_ON_CLOSE);
    positionFrame();

    setupButtons();
    setupTeilnehmerField();
  };

  public static Program getProgram() {
    if(program == null) {
      program = new Program();
    }
    return program;
  }

  public void run() {
    mainFrame.setVisible(true);
  }

  private void positionFrame() {
    Dimension windowSize = mainFrame.getSize();
    GraphicsEnvironment ge = GraphicsEnvironment.getLocalGraphicsEnvironment();
    Point centerPoint = ge.getCenterPoint();
    int dx = centerPoint.x - windowSize.width / 2;
    int dy = centerPoint.y - windowSize.height / 2;
    mainFrame.setLocation(dx, dy);
  }

  private void setupButtons() {
    mainWindowForm.getTeilnehmerSpeichernButton().setAction(new AbstractAction() {
      @Override
      public void actionPerformed(ActionEvent e) {
        try {
          FileWriter fileWriter = new FileWriter(teilnehmerFile);
          String text = mainWindowForm.getTeilnehmerTextArea().getText();
          fileWriter.write(text);
          fileWriter.close();
        } catch (IOException ex) {
          ex.printStackTrace();
        }
      }
    });
    mainWindowForm.getTeilnehmerSpeichernButton().setText("Teilnehmer speichern");

    mainWindowForm.getGenerierenButton().setAction(new AbstractAction() {
      @Override
      public void actionPerformed(ActionEvent e) {

        String[] names = mainWindowForm.getTeilnehmerTextArea().getText().split("\n");

        for(String studentName:names) {
          exportLogic.createDocument(studentName,
            mainWindowForm.getKlasseTextField().getText(),
            mainWindowForm.getTrainerTextField().getText(),
            mainWindowForm.getLernfeldTextField().getText(),
            mainWindowForm.getBerichtNrTextField().getText(),
            mainWindowForm.getJahrTextField().getText(),
            mainWindowForm.getWocheTextField().getText(),
            mainWindowForm.getMoTextArea().getText(),
            mainWindowForm.getDiTextArea().getText(),
            mainWindowForm.getMiTextArea().getText(),
            mainWindowForm.getDoTextArea().getText(),
            mainWindowForm.getFrTextArea().getText());
        }
        JOptionPane.showMessageDialog(new JFrame(), "Erfolg! " + names.length + " Berichte generiert");
      }
    });
    mainWindowForm.getGenerierenButton().setText("Berichte generieren");
  }

  private void setupTeilnehmerField() {
    try {
      BufferedReader bufferedReader = new BufferedReader(new FileReader(teilnehmerFile));
      StringBuilder stringBuilder = new StringBuilder();
      String line = bufferedReader.readLine();
      while(line != null) {
        stringBuilder.append(line).append("\n");
        line = bufferedReader.readLine();
      }
      if(!stringBuilder.toString().isEmpty()) {
        mainWindowForm.getTeilnehmerTextArea().setText(stringBuilder.toString());
      }
    } catch (FileNotFoundException e) {
      e.printStackTrace();
    } catch (IOException e) {
      e.printStackTrace();
    }
  }
}


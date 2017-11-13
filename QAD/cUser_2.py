from sys import argv
import time
import os

script, filename = argv

fl = open(filename, 'w')

question1 = raw_input("Type the full name: \n")
question2 = raw_input("Type the login now: \n")
start = "/u01/scripts/usr/_mfgadd.bash"
middle =  "/u01/scripts/usr/spool/br/"
final = "/u01/scripts/usr/spool/br/"

fl.write(question2 + ';(LIM) ' + question1 + ';135prod')
fl.write("\n")

newline = raw_input("Would you like to add new users? Please type 'y' for yes and 'n' for no \n")

while (newline == 'y'):
      question1 = raw_input("Type the full name: \n")
      question2 = raw_input("Type the login now: \n")
      fl.write(question2 + ';(LIM) ' + question1 + ';135prod')
      fl.write("\n")
      newline = raw_input("Would you like to add new users? Please type 'y' for yes and 'n' for no \n")
      if (newline == 'n'):
          break
      else:
          True
urlFinal = filename.replace(".csv","")
print "/u01/scripts/usr/_mfgadd.bash" + ' ' + "/u01/scripts/usr/spool/br/" + filename + ' > ' + "/u01/scripts/usr/spool/br/" + urlFinal + '.log'


runScript = "/u01/scripts/usr/_mfgadd.bash" + ' ' + "/u01/scripts/usr/spool/br/" + filename + ' > ' + "/u01/scripts/usr/spool/br/" + urlFinal + '.log'

print runScript

os.system(runScript)

#time.sleep(5.5)
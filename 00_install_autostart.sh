cat <<"EOF" > /tmp/install_autostart.sh
#!/bin/bash

if [ "$(/usr/bin/id -u)" -ne "0" ]; then
  echo "Root access required."
  exit 1
fi

export PATH=/usr/sbin:/usr/bin:/sbin:/bin

if grep -q '# AUTOSTART' /emulator/menu_launcher.sh; then
  echo AUTOSTART already installed.
  exit 1
fi

cp /emulator/menu_launcher.sh /userdata/menu_launcher.sh.$(date '+%Y_%m_%d_%H_%M_%S')

mount -o remount,rw /emulator

if grep -q '/emulator/atg_lcdboard' /emulator/menu_launcher.sh; then
  sed '/^\/emulator\/atg_lcdboard/i \
\
# AUTOSTART BEGIN\
if [ -d /media/usb0/autostart ]; then\
  cp -R /media/usb0/autostart /tmp/\
  chmod +x /tmp/autostart/*.sh\
  for f in `ls /tmp/autostart/*.sh`; do $f; done\
fi\
# AUTOSTART END\
' /emulator/menu_launcher.sh > /tmp/menu_launcher.sh && chmod +x /tmp/menu_launcher.sh && mv /tmp/menu_launcher.sh /emulator/menu_launcher.sh
else
  sed '/^while [ 1 -eq 1 ]/i \
\
# AUTOSTART BEGIN\
if [ -d /media/usb0/autostart ]; then\
  cp -R /media/usb0/autostart /tmp/\
  chmod +x /tmp/autostart/*.sh\
  for f in `ls /tmp/autostart/*.sh`; do $f; done\
fi\
# AUTOSTART END\
' /emulator/menu_launcher.sh > /tmp/menu_launcher.sh && chmod +x /tmp/menu_launcher.sh && mv /tmp/menu_launcher.sh /emulator/menu_launcher.sh
fi

sync

echo AUTOSTART installed
EOF

chmod +x /tmp/install_autostart.sh
$SUDO /tmp/install_autostart.sh

rm -f /tmp/install_autostart.sh


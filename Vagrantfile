# -*- mode: ruby -*-
# vi: set ft=ruby :

# Vagrantfile API/syntax version. Don't touch unless you know what you're doing!
VAGRANTFILE_API_VERSION = "2"

Vagrant.configure(VAGRANTFILE_API_VERSION) do |config|

  config.vm.box = "ubuntu/trusty32"
  
  config.vm.synced_folder "", "/home/vagrant/vagrant"

  config.vm.provision "shell", inline: "apt-get -y install python-pip"
  config.vm.provision "shell", inline: "pip install -r /home/vagrant/vagrant/requirements.txt"

  # Create a forwarded port mapping which allows access to a specific port
  # within the machine from a port on the host machine. In the example below,
  # accessing "localhost:8080" will access port 80 on the guest machine.
  # config.vm.network "forwarded_port", guest: 80, host: 8080

end

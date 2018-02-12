import paramiko
import os
files = os.listdir('./csv/')
ssh = paramiko.SSHClient()
ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
ssh.connect('host', username=your_username, password=your_password)

print "connected successfully!"

sftp = ssh.open_sftp() 
for file in files:
	sftp.put(local_file_location, remote_file_location) 
sftp.close() 
print "copied successfully!"


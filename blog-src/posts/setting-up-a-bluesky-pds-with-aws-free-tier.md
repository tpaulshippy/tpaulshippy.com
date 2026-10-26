---
title: "Setting up a Bluesky PDS with AWS Free Tier"
date: 2024-11-27
slug: setting-up-a-bluesky-pds-with-aws-free-tier
---

## Setup a Ubuntu server in AWS EC2

1. Sign up for a free AWS account.
2. Login to [AWS Console](https://aws.amazon.com/console/).
3. Select the region closest to you. Mine defaulted to Ohio.
4. Search for “EC2”.
5. Choose `Launch Instance`.
6. Give your instance a name, like “pds”.
7. Under “Application and OS Images” choose `Ubuntu`.
8. Under “Amazon Machine Image (AMI)” choose `Ubuntu 22.04` (per current version of [PDS README](https://github.com/bluesky-social/pds?tab=readme-ov-file#preparation-for-self-hosting-pds))
9. Under “Architecture” choose `64-bit (x86)` to stay in the free tier.
10. Under “Instance type” choose `t2.micro`
11. Under “Key pair (login)” click on **Create new key pair**  Give your key pair a name, like “pds-login”.
12. Choose ED25519 for the key pair type.
13. Choose `.pem` if you are on a Linux, macOS, or Windows with OpenSSH. If you use PuTTY on Windows, choose `.ppk`
14. Click `Create key pair`
15. Save the file on your machine for later use.

- Keep `Allow SSH traffic from` checked.

- Check `Allow HTTPS traffic from the internet`.

- Check `Allow HTTP traffic from the internet`.

- Under “Configure storage” change the number of GiB to `20` (per current version of [PDS README](https://github.com/bluesky-social/pds?tab=readme-ov-file#preparation-for-self-hosting-pds))

- Click `Launch Instance`

## Setting up DNS

I have a domain (tpaulshippy.com) that is already hosted elsewhere. I would like my pds server to be at pds.tpaulshippy.com so I opened my web host and entered the following settings:

### A Records

- Domain name: pds.tpaulshippy.com   IP Address: 35.90.227.79 (the Public IPv4 address of my EC2 instance)   TTL: 3600
- Domain name: *.pds.tpaulshippy.com   IP Address: 35.90.227.79 (the Public IPv4 address of my EC2 instance)   TTL: 3600

## Connecting to the instance

1. Login to [AWS Console](https://aws.amazon.com/console/).
2. Search for “EC2”.
3. Choose `Instances`
4. Click on your instance ID.
5. Click `Connect`.
6. Click `SSH client`.
7. Follow the instructions listed. Here’s how I do it on my mac:  Open a terminal.
8. Type `cd` followed by a space.
9. Type the directory on your machine where you placed the `.pem` file saved earlier.
10. Press return.
11. Type `chmod 400 "pds-login.pem"`
12. Press return.
13. Type `ssh -i "pds-login.pem" ubuntu@`
14. Paste the Public DNS name of the instance (ends in `.compute.amazonaws.com`)
15. Press return.
16. You should see a message that starts with **Welcome to Ubuntu**

## Installer on Ubuntu 20.04

1. Connect to the instance with SSH as specified above.
2. Run the following command to download the installer: `wget https://raw.githubusercontent.com/bluesky-social/pds/main/installer.sh`
3. Run the following command to run the installer: `sudo bash installer.sh`
4. Enter your domain. I put in “pds.tpaulshippy.com”. Press return.
5. Enter an email address. Press return.
6. Wait for the installation…
7. Setup the initial user as instructed. Save the password and use it to login.

## Check installation

[https://github.com/bluesky-social/pds?tab=readme-ov-file#verifying-that-your-pds-is-online-and-accessible](https://github.com/bluesky-social/pds?tab=readme-ov-file#verifying-that-your-pds-is-online-and-accessible)

## Connect to your server

[https://github.com/bluesky-social/pds?tab=readme-ov-file#using-the-bluesky-app-with-your-pds](https://github.com/bluesky-social/pds?tab=readme-ov-file#using-the-bluesky-app-with-your-pds)

## Notes

My web host did not easily allow me to add the *.pds.tpaulshippy.com DNS entry. So when I connected to my server, it showed “Invalid Handle” as my profile. But I could still browse around and see content.

I ended up switching my name servers to my domain registrar which does support wildcards on subdomains. After quite a bit of waiting for DNS propogation, I got it working.

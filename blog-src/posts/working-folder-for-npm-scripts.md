---
title: "Working folder for npm scripts"
date: 2020-07-01
slug: working-folder-for-npm-scripts
---

#### Context

I am setting up some scripts for using terraform to deploy a node project to AWS and when I execute the commands from the root folder, they do not work. They must be executed from the folder that contains the .tf files. But I am not sure how to do this in my package.json file. The scripts property looks like this:

```
"scripts": {
        "start": "node server.js",
        ...
        "deploy": "terraform apply" <-- must be run from subfolder
}
```

#### Search terms

npm scripts working directory

#### Helpful link

[https://stackoverflow.com/questions/30286498/change-working-directory-for-npm-scripts](https://stackoverflow.com/questions/30286498/change-working-directory-for-npm-scripts)

#### Outcome

Thanks to the answer provided by [eljefedelrodeodeljefe](https://stackoverflow.com/users/3580261/eljefedelrodeodeljefe), I learned that you simply have to include “cd FOLDER_NAME &&” before your script and this works across most platforms. So in my case, this looks like the following:

```
"scripts": {
        "start": "node server.js",
        ...
        "deploy": "cd infrastructure && terraform apply" <-- must be run from subfolder
}
```

One nice thing is that this approach is consistent with the syntax used in a Makefile for a Python project I have recently worked with.

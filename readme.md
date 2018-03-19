# A sample code to access calendar for Microsoft Exchange Server

## Prerequisite

* Maven

## Installation

To install, type

    $ cd my-app
    $ mvn install


## Execute

Now, to print recent appointments, type

    $ cd my-app
    $ mvn package
    $ java -cp target/my-app-1.0-SNAPSHOT.jar com.mycompany.app.App [server] [email] [password]

Running `mvn package` does a compile and creates the target directory, including a jar:


# DMSOutlookAutoArchive

Move Outlook items from mailbox to archive, e.g. PST file

## Getting Started

These instructions will get you a copy of the project up and running on your local machine for development and testing purposes. See deployment for notes on how to deploy the project on a live system.

### Prerequisites

What things you need to install the software and how to install them

```
Give examples
```

### Installing

A step by step series of examples that tell you how to get a development env running

Say what the step will be

```
Give the example
```

And repeat

```
until finished
```

End with an example of getting some data out of the system or using it for a little demo

## Authors

* **Mikhail Danshin** - *Initial work* - [mdanshin](https://github.com/mdanshin)

See also the list of [contributors](https://github.com/mdanshin/DMSOutlookAutoArchive/graphs/contributors) who participated in this project.

## License

This project is licensed under the MIT License - see the [LICENSE.md](LICENSE.md) file for details

## Acknowledgments

TBD

## Known issues
If PSSecurityException occur, try the following:

```
$cert = Get-ChildItem Cert:\LocalMachine\My\
Set-AuthenticodeSignature -Certificate $cert -FilePath "Path to file"
```
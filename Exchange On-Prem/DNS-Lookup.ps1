$Type = "mx"
$DomainNames = "jsgruppen.com","jsworldmedia.com","jssuomi.fi","triad-co.com","jsgruppe.com","groupjs.com","jssverige.se","jsdanmark.dk","jsdeutschland.de","jsoesterreich.at","jsespana.es","jsnorge.no","jsnederland.nl","jsmediatools.com","jsoutlook.com","jsosterreich.at","jsisland.is","jsitalia.it","jsbelgie.be","jsbelgien.be","jsbelgique.be","jsinternational.info","jsinternational.org","jsoesterreich.co.at","xn--jssterreich-sfb.at","jsosterreich.co.at","xn--jssterreich-sfb.co.at","jsschweiz.ch","jssuisse.ch","jssvizra.ch","jssvizzera.ch","jsuomi.fi","jsverige.se","jsintranet.com","seebrochure.com","sebrochure.dk","sebroschyr.se","unserebroschuere.at","nuestrofolleto.es","lanostrabrochure.it","visbrosjyre.no","katsoesitettae.fi","unserebroschuere.de","ziebrochure.nl","esitteemme.fi","communicationcompass.com","jscommunicationcompass.com","jsupgrade.com","jsgoto.com","jsfrance.fr","jsglobalit.com","jsgruppen.dk","notrebrochure.fr","viewbrochure.us","jsusa.us","jsamerica.us","de.ibm.com","nuestrocatalogo.es","sjabaekling.is","unserebroschuere.ch","jsengland.co.uk","jsmediacorporation.co.uk","jscomunicacion.es","jsmediacorp.co.uk","jsprofilm.dk","jsmag.jsworldmedia.com","mysite.ziebrochure.nl","cisco.jsworldmedia.com","hamburg-expressway-edge.cisco-hh.jsworldmedia.com","aarhus-expressway-edge.cisco.jsworldmedia.com","cisco-hh.jsworldmedia.com","munich-expressway-edge.cisco-mn.jsworldmedia.com","cisco-mn.jsworldmedia.com","jsprofilm.se","jsprofilm.es","jsprofilm.fi","jsprofilm.de","jsprofilm.at","jsprofilm.ch","jsprofilm.no","jsprofilm.com","jsprofilm.is"
#$DomainNames = "jsgruppen.com","jsworldmedia.com"
$ValueExport = "NameExchange"
$file = "C:\perflogs\nslookup-" + $type + ".txt"
$errorlog = "C:\perflogs\nslookup-error.txt"


Try {

#MX Record export
ForEach ($Domain in $DomainNames) 
{
            Resolve-DnsName -Type $Type -Name $Domain | Select Name,TTL,$ValueExport | export-csv -Path $file -Encoding Unicode -Append
}   

#Autodiscover cname
$ValueExport = "NameHost"
$Type = "cname"
$file = "C:\perflogs\nslookup-" + $type + ".txt"

ForEach ($Domain in $DomainNames) 
{
            Resolve-DnsName -Type $Type -Name $Domain | Select Name,TTL,$ValueExport | export-csv -Path $file -Encoding Unicode -Append
}   

#Autodiscover srv
$ValueExport = "NameTarget"
$Type = "srv"
$file = "C:\perflogs\nslookup-" + $type + ".txt"
ForEach ($Domain in $DomainNames) 
{
            $Lookup = "_autodiscover._tcp." + $Domain
            Resolve-DnsName -Type $Type -Name $Lookup | Select Name,TTL,$ValueExport | export-csv -Path $file -Encoding Unicode -Append
} 

#Autodiscover a
$ValueExport = "IPAddress"
$Type = "a"
$file = "C:\perflogs\nslookup-" + $type + ".txt"
ForEach ($Domain in $DomainNames) 
{
            $Lookup = "_autodiscover._tcp." + $Domain
            Resolve-DnsName -Type $Type -Name $Lookup | Select Name,TTL,$ValueExport | export-csv -Path $file -Encoding Unicode -Append
} 

}

catch{
       $_ | Export-Csv -Path $errorlog -Append -Encoding Unicode
    }


    # | Export-Csv -Path $errorlog -Append -Encoding Unicode
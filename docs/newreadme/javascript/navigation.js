// handle index body onload event
function load()
{
  go( 0 );
} // end function load

// handle navigation
function go( navIndex )
{
  var req = false; // ...
  var arrLinks = new Array( 7 ); // ...

  // create navigation links
  arrLinks[ 0 ] = "./html/main.html";
  arrLinks[ 1 ] = "./html/getting_started.html";
  arrLinks[ 2 ] = "./html/features.html";
  arrLinks[ 3 ] = "./xml/commands.xml";
  arrLinks[ 4 ] = "./html/custom_commands.html";
  arrLinks[ 5 ] = "./html/scripting.html";
  arrLinks[ 6 ] = "./html/advanced.html";
  arrLinks[ 7 ] = "./html/donate.html";

  // instantiate object
  req = new XMLHttpRequest();


  if ( navIndex == 3 )
  {
    req.onreadystatechange =
      function()
      {
         createCommandTable( "70", "content", req.responseXML );
      }; // end function

    // create & send file request
    req.open( "GET", arrLinks[ navIndex ], true );
    req.overrideMimeType( "text/xml" );
    req.send( null );
  }
  else
  {
    req.onreadystatechange =
      function()
      {
        document.getElementById( "content" ).innerHTML = req.responseText;
      }; // end function

    // create & send file request
    req.open( "GET", arrLinks[ navIndex ], true );
    req.send( null );
  } // end if
} // end function go

// ...
function createCommandTable( access, destination, data )
{
     var commands = data.getElementsByTagName( "command" );

     // grab div handle
     var div = document.getElementById( destination );

     // clear div content
     div.innerHTML = null;

     // create table
     var table = document.createElement( "table" );

     // create table header
     createTableHeader( table, new Array( "Command Name", "Rank", "Alias(es)" ) );

     // create table body
     var tbody = document.createElement( "tbody" );

     // loop through commands
     for ( var i = 0; i < commands.length; i++ )
     {
        var name = new String( commands[ i ].getAttribute( "name" ) ); // store name
        var token = new String( commands[ i ].getAttribute( "token" ) ); // store token
        var rank = commands[ i ].getElementsByTagName( "rank" )[ 0 ].firstChild.nodeValue; // store rank
        var aliases = commands[ i ].getElementsByTagName( "alias" ); // store aliases
        var description = commands[ i ].getElementsByTagName( "description" )[ 0 ]; // store description
        //var syntax = commands[ i ].getElementsByTagName( "syntax" ); // store syntax

        var tr = null; // store table row
        var td = null; // store table cell
        var data = null; // store table cell data
        var a = null; // store link data
        var p = null; // store paragraph data
	var row_class = new String(); // store table row class name
        
        var tmpbuf = new String(); // store temporary buffer data

        // create table row
        var tr = document.createElement( "tr" );

        tr.setAttribute( "id", name );

        // style table row
        if ( ( tbody.getElementsByTagName( "tr" ).length / 2 ) % 2 )
        {
          row_class = "evenrow";
        }
        else
        {
          row_class = "oddrow";
        } // end if

        tr.setAttribute( "class", row_class + " command" );

        /*
        ################################################
        #####         COMMAND NAME COLUMN          #####
        ################################################
        */

        var td = document.createElement( "td" );

        // is channel operator status required for this
        // command to function correctly?
        if ( token.indexOf( "operator" ) != -1 )
        {
          // create link element
          a = document.createElement( "a" );

          a.setAttribute( "href", "http://www.battle.net/info/icons.shtml" );

          // create image element
          var img = document.createElement( "img" );

          img.setAttribute( "src", "./images/gavel.gif" );
          img.setAttribute( "title", "This command requires channel operator status." );
          
          // add image to link
          a.appendChild( img );

          // add link to cell
          td.appendChild( a );
        } // end if
 
        a = document.createElement( "a" );

        a.setAttribute( "href", "javascript:showCommand('" + name + "')" );

        data = document.createTextNode( name );

        a.appendChild( data );

        td.appendChild( a );  // add data to cell
        tr.appendChild( td ); // add cell to row

        /*
        ################################################
        #####         COMMAND RANK COLUMN          #####
        ################################################
        */

        td = document.createElement( "td" );
        data = document.createTextNode( rank );

        td.appendChild( data );
        tr.appendChild( td );

        /*
        ################################################
        #####       COMMAND ALIASES COLUMN         #####
        ################################################
        */            

        if ( aliases.length > 0 )
        {
          for ( var j = 0; j < aliases.length; j++ )
          {
            var alias = aliases[ j ].firstChild.nodeValue; // store alias name

            tmpbuf += alias;

            if ( j < aliases.length - 1 )           
            {
              tmpbuf += ", ";
            } // end if
          } // end loop
        }
        else
        {
          tmpbuf = "-";
        } // end if

        td = document.createElement( "td" );
        data = document.createTextNode( tmpbuf );

        td.appendChild( data );
        tr.appendChild( td );

        tbody.appendChild( tr ); // add row to table body

        /*
        ################################################
        #####      COMMAND DOCUMENTATION ROW       #####
        ################################################
        */

        tr = document.createElement( "tr" );
        
        tr.setAttribute( "style", "display:none" );
	tr.setAttribute( "class", row_class + " description" );

        td = document.createElement( "td" );
        
        td.setAttribute( "colspan", "3" );

        p = document.createElement( "p" )
        
        data = document.createTextNode( description.firstChild.nodeValue );

        p.appendChild( data );

        /*
        if ( syntax.length > 0 )
        {
        p.appendChild( document.createElement( "br" ) );
        p.appendChild( document.createElement( "br" ) );
        
        data = document.createElement( "strong" );
        data.appendChild( document.createTextNode( "Syntax: " ) );

        p.appendChild( data );

        data = document.createTextNode( syntax[ 0 ].firstChild.nodeValue );

        p.appendChild( data );
        }
        */

        td.appendChild( p );
        tr.appendChild( td )

        tbody.appendChild( tr ); // add row to table body
     } // end loop

     // add table body to table
     table.appendChild( tbody );

     // add table to content div
     div.appendChild( table );
} // end function createCommandTable

// ...
function createTableHeader( table, columns )
{
  var thead = document.createElement( "thead" );
  var tr = null;
  var td = null;
  var data = null;

  tr = document.createElement( "tr" );

  for( var i = 0; i < columns.length; i++ )
  {
    td = document.createElement( "td" );
    data = document.createTextNode( columns[ i ] );

    td.appendChild( data );
    tr.appendChild( td );
  }
  
  thead.appendChild( tr );

  table.appendChild( thead );
} // end function createTableHeader

// ...
function showCommand( command )
{
  var row = null; // ...

  // expand command description

  row = document.getElementById( command ).nextSibling;

  if ( row.getAttribute( "style" ) == 'display: table-row;' )
  {
    row.setAttribute( "style", "display: none" );
  }
  else
  {
    row.setAttribute( "style", "display: table-row" );
  } // end if

  // collapse unrelated command descriptions

  row = document.getElementsByTagName( "tr" );

  for ( var i = 0; i < row.length; i++ )
  {
    var row_id = new String( row[ i ].getAttribute( "id" ) );
    var row_class = new String( row[ i ].getAttribute( "class" ) );

    if ( row_class.indexOf( "command" ) != -1 )
    {
      if ( row_id != command )
      {
        row[ i ].nextSibling.setAttribute( "style", "display: none" );
      } // end if
    } // end if
  } // end loop
} // end function showCommand
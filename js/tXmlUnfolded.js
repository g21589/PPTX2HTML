/**
 * @author: Tobias Nickel
 * @date: 06.04.2015
 * I needed a small xmlparser chat can be used in a worker. 
 */

/**
 * parseXML / html into a DOM Object. with no validation and some failur tolerance
 * @params S {string} your XML to parse
 */ 
function tXml(S){
    "use strict";
    var openBracket = "<";
    var openBracketCC = "<".charCodeAt(0);
    var closeBracket = ">";
    var closeBracketCC = ">".charCodeAt(0);
    var minus = "-";
    var minusCC = "-".charCodeAt(0);
    var slash = "/";
    var slashCC = "/".charCodeAt(0);
	var exclamation = '!';
	var exclamationCC = '!'.charCodeAt(0);
	var singleQuote = "'";
	var singleQuoteCC = "'".charCodeAt(0);
	var doubleQuote = '"';
	var doubleQuoteCC = '"'.charCodeAt(0);
    
    /**
     * parsing a list of entries
     */
    function parseChildren(){
        var children = [];
        while(S[pos]){
            if(S.charCodeAt(pos) == openBracketCC){
                if(S.charCodeAt(pos+1) === slashCC){
                    //while(S[pos]!=='>'){ pos++; }
                    pos = S.indexOf(closeBracket,pos);
                    return children;
                }else if(S.charCodeAt(pos+1) === exclamationCC){
                    if(S.charCodeAt(pos+2) == minusCC){
						//comment support
                        while(!(S.charCodeAt(pos)===closeBracketCC && S.charCodeAt(pos-1)==minusCC && S.charCodeAt(pos-2)==minusCC &&pos != -1)){pos = S.indexOf(closeBracket, pos+1);}
                        if(pos===-1) pos=S.length
                    }else{
						// doctypesupport
                        pos+=2;
                        while(S.charCodeAt(pos)!==closeBracketCC){ pos++; }
                    }
                    pos++;
                    continue;
                }
				var node = {};
				pos++;
				var startNamePos = pos;
				while(nameSpacer.indexOf(S[pos])===-1){ pos++; }
				var node_tagName = S.slice(startNamePos,pos);

				// parsing attributes
				var attrFound=false;
				while (S.charCodeAt(pos) !== closeBracketCC) {
					var c = S.charCodeAt(pos);
					if ((c>64&&c<91)||(c>96&&c<123)) {
					//if('abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ'.indexOf(S[pos])!==-1 ){
					
						startNamePos = pos;
						while (nameSpacer.indexOf(S[pos])===-1){
							pos++;
						}
						var name = S.slice(startNamePos,pos);
						// search beginning of the string
						var code=S.charCodeAt(pos);
						while(code !== singleQuoteCC && code !== doubleQuoteCC){ pos++;code=S.charCodeAt(pos) }
						
						
						var startChar = S[pos];
						var startStringPos= ++pos;
						pos = S.indexOf(startChar,startStringPos);
						var value = S.slice(startStringPos,pos);
						if(!attrFound){
							var node_attributes = {};
							attrFound=true;
						}
						node_attributes[name] = value;
					}
					pos++;

				}
				// optional parsing of children
				if (S.charCodeAt(pos-1) !== slashCC) { 
					if (node.tagName == "script") {
						var start=pos;
						pos=S.indexOf('</script>',pos);
						node.children=[S.slice(start,pos-1)];
						pos+=8;
					} else if (node_tagName == "style") {
						var start=pos;
						pos=S.indexOf('</style>',pos);
						node.children=[S.slice(start,pos-1)];
						pos+=7;
					} else if (!NoChildNodes[node.tagName]) {
						pos++;
						var node_children = parseChildren(name);
					}
				}
                children.push({
					"children": node_children,
					"tagName": node_tagName,
					"attrs": node_attributes
				});
            } else {
				var startTextPos = pos;
				pos = S.indexOf(openBracket,pos)-1;
				if(pos===-2)pos=S.length;
                var text = S.slice(startTextPos,pos+1);
                if(text.trim().length>0) {
					children.push(text);
				}
            }
            pos++;
        }
        return children;
    }
    
	/**
     *    returns text until the first nonAlphebetic letter
     */
    var nameSpacer = '\n\t>/= ';
	
    /**
     *    is parsing a node, including tagName, Attributes and its children,
     * to parse children it uses the parseChildren again, that makes the parsing recursive
     */
    var NoChildNodes={};
    
	var pos=0;
    return simplefy(parseChildren());
}

function simplefy(children) {
    var out = {};
    
	if (children === undefined) {
		return {};
	}
	
    if (children.length === 1 && typeof children[0] == 'string') {
        return children[0];
	}

    // map each object
    children.forEach(function(child) {

        if (!out[child.tagName]) {
            out[child.tagName] = [];
		}
		
        if (typeof child == 'object') {
            var kids = simplefy(child.children);
            out[child.tagName].push(kids);
            if (child.attrs) {
                kids.attrs = child.attrs;
            }
        } else {
            out[child.tagName].push(child);
        }
		
    });
    
    for (var i in out) {
        if (out[i].length == 1) {
            out[i] = out[i][0];
        }
    }
    
    return out;
};

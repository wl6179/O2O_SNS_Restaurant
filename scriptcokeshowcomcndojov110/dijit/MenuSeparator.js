/*
	Copyright (c) 2004-2009, The Dojo Foundation All Rights Reserved.
	Available via Academic Free License >= 2.1 OR the modified BSD license.
	see: http://dojotoolkit.org/license for details
*/


if(!dojo._hasResource["dijit.MenuSeparator"]){dojo._hasResource["dijit.MenuSeparator"]=true;dojo.provide("dijit.MenuSeparator");dojo.require("dijit._Widget");dojo.require("dijit._Templated");dojo.require("dijit._Contained");dojo.declare("dijit.MenuSeparator",[dijit._Widget,dijit._Templated,dijit._Contained],{templateString:"<tr class=\"dijitMenuSeparator\">\r\n\t<td colspan=\"4\">\r\n\t\t<div class=\"dijitMenuSeparatorTop\"></div>\r\n\t\t<div class=\"dijitMenuSeparatorBottom\"></div>\r\n\t</td>\r\n</tr>\r\n",postCreate:function(){dojo.setSelectable(this.domNode,false);},isFocusable:function(){return false;}});}
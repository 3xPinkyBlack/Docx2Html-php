<?php
/* 
	THIS PHP CLASS ALLOWS A USER TO CONVERT A DOCX FILE TO AN HTML FORMAT
	at this time this class only supports the following HTML Components
	1. table with colspan and rowspans
	2. headings(h1, h2, h3, h4, h5, h6)
	3. colors
	4. images
	5. superscript
	6. subscript
	7. lists(numbered, disc)
	8. links
	
	i used jquery to remove lists with empty value.

	i am working on other features and i will update this class if the feature works

	if you want any feature to be added in this class please pull a request in my github
	repository.
    https://github.com/3xpinkyblack
	any issue contact me.

	Created By: Habib Endris
	Creator Website: habib-kiot.tk
	Creator Email Address: rasorhabib@gmail.com / 3xpinkyblack@gmail.com

	this class is for personal use only for commercial use please contact me with the above
	email address.
*/ 

class Doc2Txt {
	private $docxFileName;
	
	public function __construct($filePath) {
		$this->docxFileName = $filePath;
	}

	public function read_docx(){
		$content = "";
		$resources = "";
		$medias = NULL;
		$numberings = "";

		$zip = new ZipArchive();
		if($zip->open($this->docxFileName)) {
			for($i = 0; $i < $zip->numFiles; $i++) {
				$name = $zip->getNameIndex($i);
				if($name == "word/document.xml") {
					$fp = $zip->getStream($name);
					while(!feof($fp)) {
						$content .= fread($fp, 2);
					}
				}

				if($name == "word/_rels/document.xml.rels") {
					$fp = $zip->getStream($name);
					while(!feof($fp)) {
						$resources .= fread($fp, 2);
					}
				}

				if($name == "word/numbering.xml") {
					$fp = $zip->getStream($name);
					while(!feof($fp)) {
						$numberings .= fread($fp, 2);
					}
				}

				if(substr($name,0,11) == "word/media/") {
					$fp = $zip->getStream($name);
					$zip->extractTo("medias",$name);
					$medias .= $name;
				}
			}
		}


		return $this->findTags($content, $resources, $numberings);
	}

	public function extractList($numberings) {
		
	}

	private function findTags($content, $resources, $numberings) {
		$tags_with_newline = "";
		//this code alligns all the tags to a new line
		for($x = 0; $x < strlen($content); $x++) {
			$let = $content[$x];
			if($let == "<") {
				$let = "\n<";
			}

			$tags_with_newline .= $let;
		}

		//this code remove new line with no values
		$tags_with_newline = explode("\n",$tags_with_newline);
		$content_without_null_line = array();
		foreach($tags_with_newline as $str) {
			if(trim($str) != "") {
				if(strchr($str, "gridSpan") == TRUE) {
					explode("\"", $str)[1];
					for($a = 0; $a < explode("\"", $str)[1]; $a++) {
						// array_push($content_without_null_line, "<w:tc id=\"novalue\">");
						// array_push($content_without_null_line, "</w:tc>");
					}
				}
				array_push($content_without_null_line, $str);
			}
		}

		//this line finds the opening and closing tags
		$main_tags = array(
			"w:b","w:body","w:p","w:r","w:t","w:pPr","w:rPr","w:sectPr","w:bCs","w:i","w:iCs",
			"w:pStyle","w:u","w:rFonts","w:type","w:docGrid","w:pgSz","w:pgMar","w:pgNumType",
			"w:formProt","w:textDirection","w:sz","w:szCs","w:tab","w:caps","w:smallCaps",
			"w:spacing","w:color","w:highlight","w:ind","w:hanging","w:pageBreakBefore","w:jc",
			"w:firstLine","w:wordWrap","w:pgBorders","w:dstrike","w:sz","w:strike","w:vertAlign","w:shd",
			"w:pBdr","w:hyperlink","w:rStyle","w:drawing","wp:anchor","wp:simplePos","wp:positionH",
			"wp:posOffset","wp:extent","wp:effectExtent","wp:wrapSquare","docPr","wp:positionV",
			"wp:cNvGraphicFramePr","a:graphic","a:graphicData","a:graphicFrameLocks","pic:pic",
			"pic:nvPicPr","pic:cNvPr","a:picLocks","a:blip","pic:blipFill","a:stretch","a:fillRect",
			"a:xfrm","a:off","a:ext","a:prstGeom","pic:spPr","a:avLst","w:numId","w:numPr","wp:align",
			"w:tbl","w:tblPr","w:tblW","w:tblInd","w:tblBorders","w:top","w:left","w:bottom","w:insideH",
			"w:right","tblCellMar","w:tblGrid","w:gridCol","w:tr","w:trPr","w:tc","w:tcPr","w:tcW",
			"w:insideV","w:vMerge","w:gridSpan","w:trHeight"
		);

		$tags_prop = array(
			"t","b","r","l","w:val","w:ascii","xml:space","w:w","w:h","w:left","w:right","w:header",
			"w:top","w:shd","w:footer","w:bottom","w:left","w:start","w:end","w:sz","w:color","w:space",
			"w:on","w:after","w:before","w:hanging","w:firstLine","cy","cx"
		);

		$sty_tags = array("w:rFonts","w:b","w:bCs","w:i","w:iCs","w:color","w:u","w:highlight","w:pgSz","w:pgMar",
						"w:pgBorders","w:dstrike","w:sz","w:strike","w:smallCaps","w:caps","w:vertAlign","w:shd",
						"w:pBdr","w:spacing","w:ind","w:hanging","w:pageBreakBefore","w:jc","w:firstLine","w:wordWrap",
						"w:pgNumType","w:formProt","w:textDirection","w:docGrid","w:type","w:pStyle","w:tab","w:rStyle",
						"wp:simplePos","wp:effectExtent","wp:wrapSquare","wp:extent","a:picLocks","a:fillReact",
						"a:off","a:ext","a:avLst","a:fillRect","a:graphicFrameLocks","w:numId","w:tblW","w:tblInd",
						"w:top","w:left","w:bottom","w:insideH","w:right","w:gridCol","w:tcW","w:insideV",
						"w:vMerge","w:gridSpan","w:trHeight"
				);

		$tags_rep = array(
			"w:rFonts" => array("font-family: sans-serif;",
							array("w:ascii" => "font-family: ","w:cs" => "font-family: ")),
			"w:b" => array("font-weight: bold;",array("w:val" => "font-weight: ")),
			"w:bCs" => array("font-weight: bold;",array("w:val" => "font-weight: ")),
			"w:i" => array("font-style: italic;", array("w:val" => "font-style: ")),
			"w:iCs" => array("font-style: italic;",array("w:val" => "font-style: ")),
			"w:color" => array("color: black;", array("w:val" => "color: ")),
			"w:u" => array("text-decoration: none;", array("w:val" => "text-decoration-style: ")),
			"w:highlight" => array("background-color: none", array("w:val" => "background-color: ")),
			"w:pgSz" => array("max-width: none;",array("w:w" => "max:width: ")),
			"w:pgMar" => array("margin: 0px;",
							array("w:top" => "margin-top: ", "w:bottom" => "margin-bottom: ",
								"w:left" => "margin-left: ", "w:right" => "margin-right: "
							)),
			"w:pgBorders" => array("border: 0px;",
								array("w:top" => "border-top: ", "w:bottom" => "border-bottom: ",
									"w:left" => "border-left: ", "w:right" => "border-right: ",
									"w:color" => "border-color: ", "w:sz" => "border-width: ",
									"w:val" => "border-style: ", "w:space" => "border-spacing: "
								)),
			"w:dstrike" => array("text-decoration-style: none;", array("w:val" => "text-decoration-style: ")),
			"w:sz" => array("font-size: auto;", array("w:val" => "font-size: ")),
			"w:szCs" => array("font-size: auto;", array("w:val" => "font-size: ")),
			"w:strike" => array("text-decoration: line-through;", array("w:on" => "text-decoration: ")),
			"w:smallCaps" => array("text-transform: uppercase; font-size: small;",
								 array("w:val" => "font-size: small; text-transform: ")),
			"w:caps" => array("text-transform: uppercase;", array("w:val" => "text-transform: ")),
			"w:vertAlign" => array("vertical-align: sub;", array("w:val" => "vertical-align: ")),
			"w:shd" => array("background-color: white;", array("w:shd" => "background-color: ")),
			"w:pBdr" => array("border: 0px;",
							array("w:top" => "border-top: ", "w:bottom" => "border-bottom: ",
								"w:left" => "border-left: ", "w:right" => "border-right: ",
								"w:color" => "border-color: ", "w:sz" => "border-width: ",
								"w:val" => "border-style: ", "w:space" => "border-spacing: "
							)),
			"w:spacing" => array("margin: 0px;",
								array("w:left" => "margin-left: ", "w:start" => "margin-left: ",
									"w:right" => "margin-right: ", "w:end" => "margin-right: ",
									"w:after" => "margin-bottom: ", "w:before" => "margin-top: ")),
			"w:ind" => array("margin: 0px;",
							array("w:left" => "margin-left: ", "w:start" => "margin-left: ",
								"w:right" => "margin-right: ", "w:end" => "margin-right: ",
								"w:after" => "margin-bottom: ", "w:before" => "margin-top: ")),
			"w:hanging" => array("padding-left: 0px;", array("w:hanging" => "padding-left: ")),
			"w:pageBreakBefore" => array("page-break-before: avoid;", array("w:val" => "page-break-always: ")),
			"w:jc" => array("text-align: left;", array("w:val" => "text-align: ")),
			"w:firstLine" => array("text-indent: unset;", array("w:firstLine" => "text-indent: ")),
			"w:wordWrap" => array("vertical-align: sub;", array("w:val" => "vertical-align: ")),
			"w:pgNumType" => array(" ", array("" => "")),
			"w:formProt" => array(" ", array("" => "")),
			"w:textDirection" => array(" ", array("" => "")),
			"w:docGrid" => array(" ", array("" => "")),
			"w:type" => array(" ", array("" => "")),
			"w:pStyle" => array(" ", array("w:val" => "Heading")),
			"w:tab" => array(" ", array("" => "")),
			"w:rStyle" => array(" ", array("" => "")),
			"w:drawing" => array(" ", array("" => "")),
			"wp:anchor" => array(" ", array("" => "")),
			"wp:simplePos" => array(" ", array("" => "")),
			"wp:positionH" => array(" ", array("" => "")),
			"wp:positionV" => array(" ", array("" => "")),
			"wp:posOffset" => array(" ", array("" => "")),
			"a:picLocks" => array(" ", array("" => "")),
			"a:fillRect" => array(" ", array("" => "")),
			"a:off" => array(" ", array("" => "")),
			"a:ext" => array(" ", array("" => "")),
			"a:avLst" => array(" ", array("" => "")),
			"a:fillReact" => array(" ", array("" => "")),
			"a:graphicFrameLocks" => array(" ", array("" => "")),
			"w:numId" => array("list-style-type: none;", array("w:val" => "list-style-type: ")),
			"wp:extent" => array("height: auto; \n\t\twidth: auto;",
							 array("cy" => "height: ", "cx" => "width: ")),
			"wp:effectExtent" => array("margin: auto;", 
									array("l" => "margin-left: ", "r" => "margin-right: ",
										"b" => "margin-bottom: ", "t" => "margin-top: ")),
			"wp:wrapSquare" => array(" ", array("" => "")),
			"w:tblW" => array(" ", array("w:w" => "width: ", "w:type" => "transform-type: ")),
			"w:tblInd" => array("margin: none",
							 array("w:w" => "margin-left: ", "w:bottomFromText" => "margin-bottom: ",
								"w:topFromText" => "margin-top: ")),
			"w:top" => array(" ", array("w:val" => "border-top: ", "w:color" => "border-top-color: ",
									"w:w" => "padding-top:")),
			"w:left" => array(" ", array("w:val" => "border-left: ", "w:color" => "border-left-color: ",
									"w:w" => "padding-left: ")),
			"w:bottom" => array(" ", array("w:val" => "border-bottom: ", "w:color" => "border-bottom-color: ",
									"w:w" => "padding-bottom: ")),
			"w:insideH" => array(" ", array("w:val" => "border-right: ", "w:color" => "border-right-color: ",
									"w:w" => "padding-right: ")),
			"w:right" => array(" ", array("w:val" => "border-right: ", "w:color" => "border-right-color: ",
									"w:w" => "padding-right: ")),
			"w:gridCol" => array(" ", array("" => "")),
			"w:tcW" => array(" ", array("" => "")),
			"w:insideV" => array(" ", array("" => "")),
			"w:tblBorders" => array("border: none;",
								 array("w:top" => "border-top: ", "w:right" => "border-right: ",
									"w:bottom" => "border-bottom: ", "w:left" => "border-left: ")),
			"w:vMerge" => array(" ", array("w:val" => "row-span: ")),
			"w:gridSpan" => array(" ", array("w:val" => "column-span: ")),
			"w:trHeight" => array(" ", array("w:val" => "height: "))
		);

		$head_tags = array("");

		$styles = array();

		$live_tags = array();
		$num_tabs = 0;
		$docx_html = "docx-html";
		$latest_id = -1;
		$list_id = array();
		$list_id_type = array();

		$header_id = array();
		$header_id_type = array();

		$table_id = array();
		$table_row = array();
		$table_column = array();
		
		$row_spans = array(array());
		$row_span_counter = -1;
		$row_span_cur_counter = 0;
		$removable_cols = array();
		$col_spans = array(array());
		$col_span_counter = -1;

		$rows = 0;
		$cols = 0;
		$tables = 0;

		$restart = array();
		$continue = array();
		//iterate throuch all the array elements
		for($x = 0; $x < count($content_without_null_line); $x++) {
			$val = $content_without_null_line[$x];
			if(substr($val,0,2) == "<?" || substr($val,0,11) == "<w:document"
				|| substr($val,0,12) == "</w:document") { //check xml header or <:document tag
				$content_without_null_line[$x] = "";
			
			//check if the first word is an opening tag
			} else if(substr($val,0,3) == "<w:" || substr($val,0,4) == "<wp:" ||
					  substr($val,0,3) == "<a:" || substr($val,0,5) == "<pic:") { //check opening tag
				$tag_name = "";
				$is_spec = FALSE;
				
				//find the exact tag name 
				foreach($main_tags as $tg) {
					if(!strchr($val,"<".$tg) == FALSE) {
						$tag_name = $tg;
					}
				}

				//check if the tag found above is a style tag or not
				foreach($sty_tags as $tg) {
					if($tag_name == $tg) {
						$is_spec = TRUE;
						break;
					}
				}

				//if its a style tag
				if($is_spec == TRUE) {
					$spaces = "";
					//this line is for the indentation purpose
					for($i = 0; $i < $num_tabs; $i++) {
						$spaces .= "\t";
					} 

					$prop_find = FALSE;
					$style_property  = "";
					$props = explode("\"",$val);
					$sty_name = "";
					// iterate through the tag properties
					for($k = 0; $k < count($props); $k++) {
						if($k % 2 == 0) {
							for($j = 0; $j < count($tags_prop); $j++) {
								if(strchr($props[$k], $tags_prop[$j]."=") != FALSE) {
									// $sty_name = $tags_rep[$tag_name][1][0];
									if(isset($tags_rep[$tag_name][1][$tags_prop[$j]])) {
										$sty_name = $tags_rep[$tag_name][1][$tags_prop[$j]];
										
										$prop_find = TRUE;
										if(strchr($val, "=\"".$props[$k+1]."\"") && $sty_name != "" &&
										!strchr($style_property, $sty_name.str_replace(";",", ", $props[$k+1]))) {
											$prop_value = $props[$k+1];
											if($tag_name == "w:gridSpan") {
												$id = trim(chop(explode("\n",$styles[$latest_id])[1],"{"));
												$id = ltrim($id,"#");
												$col_span_counter++;
												$col_spans[$col_span_counter] = $id;
												$col_spans[$col_span_counter][0] = $prop_value;
												$cols += $prop_value - 1;
											} else if($tag_name == "w:vMerge") {
												if($prop_value == "restart") {
													$id = trim(chop(explode("\n",$styles[$latest_id])[1],"{"));
													$id = ltrim($id,"#");
													$row_span_counter++;
													$row_spans[$row_span_counter] = "d".$id;
													$row_spans[$row_span_counter][0] = 1;
													$row_span_cur_counter = 1;
													array_push($restart, $cols);
												} else {
													$id = trim(chop(explode("\n",$styles[$latest_id])[1],"{"));
													$id = ltrim($id,"#");
													array_push($removable_cols,$id);
													$row_span_cur_counter++;
													$row_spans[$row_span_counter][0] = $row_span_cur_counter;
													array_push($continue, $cols);
												}
											} else if($tag_name == "wp:extent") {
												$val_num = (int)((int)$prop_value / 9000);
												$style_property .= "\t\t\t".$sty_name.$val_num.";\n";
												$sty_name = "";
											} else if($tag_name == "w:sz" || $tag_name == "w:szCs") {
												$val_num = (int)((int)$prop_value / (2 - 1/2));
												$style_property .= "\t\t\t".$sty_name.$val_num.";\n";
												$sty_name = "";
											} else if($tag_name == "w:numId") {
												$img_id = trim(chop(explode("\n",$styles[$latest_id])[1],"{"));
												$img_id = ltrim($img_id,"#");

												array_push($list_id, $img_id);
												array_push($list_id_type, $prop_value);
												// $repl = array(0 => "disc", 1 => "decimal",
												// 	2 => "lower-alpha", 3 => "lower-roman",
												// 	4 => "upper-alpha", 5 => "upper-roman"
												// );
												// $val_num = (int)$prop_value;
												// $style_property .= "\t\t".$sty_name.$repl[$val_num].";\n";
												// $sty_name = "";
											} else if($tag_name == "w:pStyle") {
												$id = trim(chop(explode("\n",$styles[$latest_id])[1],"{"));
												$id = ltrim($id,"#");

												if(trim($prop_value) != "Normal") {
													$head = substr($prop_value, 7, strlen($prop_value));
													if($head == "1" || $head == "2" || $head == "3" || $head == "4"
														|| $head == "5" || $head == "6") {
														array_push($header_id, $id);
														array_push($header_id_type, "h".$head);
													} else {
														array_push($header_id, $id);
														array_push($header_id_type, "h1");
													}
												}
											} else {
												$style_property .= "\t\t\t".$sty_name.str_replace(";",", ", $props[$k+1]).";\n";
												$sty_name = "";
											}
										}
										break;
									}
								}
							}
						}
					}

					if($prop_find == FALSE && !strchr($styles[$latest_id], $tags_rep[$tag_name][0])) {
						$style_property = "\t\t\t".$tags_rep[$tag_name][0]."\n"; 
					}

					$styles[$latest_id] .= $style_property;
					
					// $content_without_null_line[$x] = "";
					$content_without_null_line[$x] = $spaces.$val;
					// $styles[$latest_id] .= "\t\t".$tag_name.";\n";
				} else {
					$spaces = "";
					for($i = 0; $i < $num_tabs; $i++) {
						$spaces .= "\t";
					}

					if($tag_name == "w:p" || $tag_name == "w:r" || $tag_name == "w:hyperlink"
					 || $tag_name == "w:sectPr" || $tag_name == "w:t" || $tag_name == "w:drawing"
					 || $tag_name == "a:blip" || strchr($val,"w:tbl>") == TRUE ||
					 $tag_name == "w:tr" || strchr($val,"w:tc>") == TRUE) {
						if($tag_name == "w:hyperlink") {
							$link_id = "";
							$hy_prop = explode("\"",$val);
							for($y = 0; $y < count($hy_prop); $y++) {
								if($y % 2 == 0 && strchr($hy_prop[$y], "r:id") == TRUE) {
									$link_id = $hy_prop[$y+1];
									break;
								}
							}

							$hy_prop_res = explode("\"",$resources);
							for($z = 0; $z < count($hy_prop_res); $z++) {
								if(strchr($hy_prop_res[$z],"$link_id")) {
									$latest_id += 1;
									array_push($styles,"\n\t\t#".$docx_html."-".$x." {\n");
									$val = "/link/ href=\"".$hy_prop_res[$z+4]."\" target=\"_blank\" id=\"".$docx_html."-".$x."\" /link-/";
									break;
								}
							}
						} else if($tag_name == "a:blip") {
							$link_id = "";
							$hy_prop = explode("\"",$val);
							for($y = 0; $y < count($hy_prop); $y++) {
								if($y % 2 == 0 && strchr($hy_prop[$y], "r:embed") == TRUE) {
									$link_id = $hy_prop[$y+1];
									break;
								}
							}

							$hy_prop_res = explode("\"",$resources);
							for($z = 0; $z < count($hy_prop_res); $z++) {
								if(strchr($hy_prop_res[$z],"$link_id")) {
									$img_id = trim(chop(explode("\n",$styles[$latest_id])[1],"{"));
									$img_id = ltrim($img_id,"#");
									
									
									$img_width = trim(explode("\n",$styles[$latest_id])[2]);
									$img_height = trim(explode("\n",$styles[$latest_id])[3]);

									$img_height = str_replace(": ","=\"",$img_height);
									$img_width = str_replace(": ","=\"",$img_width);

									$img_height = str_replace(";","\"",$img_height);
									$img_width = str_replace(";","\"",$img_width);

									$new_img_id = "id=\"$img_id\" $img_height $img_width src=\"medias/word/".$hy_prop_res[$z+4]."\"";
									for($a = 0; $a < $x; $a++) {
										if(strchr(trim($content_without_null_line[$a]), "id=\"".$img_id."\"")) {
											$content_without_null_line[$a] = str_replace("id=\"".$img_id."\"",$new_img_id, $content_without_null_line[$a]);
										}
									}
									break;
								}
							}
						} else {
							$latest_id += 1;
							array_push($styles,"\n\t\t#".$docx_html."-".$x." {\n");
							$val = str_replace($tag_name,$tag_name." id=\"".$docx_html."-".$x."\"",$val);
							
							if(strchr($val,"w:tbl id") == TRUE) {
								$id = trim(chop(explode("\n",$styles[$latest_id])[1],"{"));
								$id = ltrim($id,"#");
								array_push($table_id, $id);
								$tables++;
							} else if(strchr($val,"w:tr id") == TRUE) {
								$id = trim(chop(explode("\n",$styles[$latest_id])[1],"{"));
								$id = ltrim($id,"#");
								array_push($table_row, $id);
								$rows++;
							} else if(strchr($val,"w:tc id") == TRUE) {
								$id = trim(chop(explode("\n",$styles[$latest_id])[1],"{"));
								$id = ltrim($id,"#");
								array_push($table_column, $id);
								$cols++;
							}
						}
					}

					$content_without_null_line[$x] = $spaces.$val;
					array_push($live_tags, $tag_name);
					$num_tabs++;
				}
			} else if(substr($val,0,2) == "</") { //check closing tag
				$tag_name = "";
				$is_spec = FALSE;
				foreach($main_tags as $tg) {
					if($tag_name == $tg) {
						$tag_name = $tg;
					}
				}

				$spaces = "";
				array_pop($live_tags);
				$num_tabs--;
				for($i = 0; $i < $num_tabs; $i++) {
					$spaces .= "\t";
				}

				if(substr($val,0,7) != "</body>") {
					$content_without_null_line[$x] = $spaces.$val;
				}
			}
		}
		
		for($j = 0; $j < count($styles); $j++) {
			$styles[$j] .= "\t\t}";
			$sty_change = explode("\n",$styles[$j]);
			$styles[$j] = NULL;
			for($e = 0; $e < count($sty_change); $e++) {
				if(strchr($sty_change[$e], "color: ") == TRUE) {
					$is_found = FALSE;
					for($g = 0; $g < 9; $g++) {
						if(substr_count($sty_change[$e], $g) != 0) {
							$is_found = TRUE;
							break;
						}
					}
					if($is_found == TRUE) {
						$styles[$j] .= str_replace("color: ","color: #",$sty_change[$e])."\n";
					} else {
						$styles[$j] .= $sty_change[$e]."\n";
					}
				} else if(strchr($sty_change[$e], "font-size: ") == TRUE) {
					$styles[$j] .= str_replace(";","px;",$sty_change[$e])."\n";
				} else if(count($sty_change) > 3){
					$styles[$j] .= $sty_change[$e]."\n";
				}
			}
		}

		$styles[0] = "\n\t<style>".$styles[0];
		$styles[count($styles) - 1] .= "\n\t</style>";

		$new_numbering = NULL;
		for($x = 0; $x < strlen($numberings); $x++) {
			$let = $numberings[$x];
			if($let == "<") {
				$let = "\n<";
			}

			$new_numbering .= $let;
		}

		// the below code adds columns for colspans
		$num_cols = $cols / $rows;

		$span_range = array(array());
		$range_counter = 0;
		for($g = 0; $g < count($restart); $g++) {
			$col_num = (int) $restart[$g] % (int) $num_cols;
			$has_range = FALSE;
			for($h = $g + 1; $h < count($restart); $h++) {
				if($col_num == (int) $restart[$h] % (int) $num_cols) {
					$span_range[$range_counter] = $restart[$g];
					$span_range[$restart[$g]][0] = $restart[$h];
					$has_range = TRUE;
				}
			}

			if($has_range == FALSE) {
				$span_range[$range_counter] = $restart[$g];
				$span_range[$restart[$g]][0] = $cols;
			}
		}

		$final_row_span = array();
		for($i = 0; $i < count($restart); $i++) {
			$col_num = (int) $restart[$i] % (int) $num_cols;
			$final_row_span[$i] = $restart[$i];
			$final_row_span[$restart[$i]][0] = 1;
			for($j = 0; $j < count($continue); $j++) {
				$col_num == (int) $continue[$j] % (int) $num_cols;
				if($restart[$i] > $continue[$j]) {
					continue;
				}

				if($continue[$j] > $span_range[$restart[$i]][0]) {
					break;
				}
				
				if($col_num == (int) $continue[$j] % (int) $num_cols) {
					$final_row_span[$restart[$i]][0] = $final_row_span[$restart[$i]][0] + 1;
				}
			}
		}

		$numberings = array();
		$new_numbering = explode("\n",$new_numbering);

		$check_dup = 0;
		$arr_counter = 0;
		for($h = 0; $h < count($new_numbering); $h++) {
			if(substr($new_numbering[$h], 0,31) == "<w:abstractNum w:abstractNumId="
			 || substr($new_numbering[$h], 0,9) == "<w:numFmt") {
				 if(substr($new_numbering[$h], 0,14) == "<w:abstractNum") {
					// array_push($numberings, $new_numbering[$h]);
					$check_dup = 0;
				 } else if($check_dup == 0) {
					$check_dup = 1;
					$new_num = explode("\"",$new_numbering[$h]);
					$numeric = 0;
					if($new_num[1] == "bullet") {
						array_push($numberings, $numeric);
						$numeric = 0;
					} else if($new_num[1] == "decimal"){
						$numeric = 1;
						array_push($numberings, $numeric);
					}
				}
			}
		}

		//the below code adds table row for colspans

		$is_table_found = FALSE;
		$table_counter = 0;
		$is_removable = FALSE;
		//the below code will structure the table
		$column_counter = 0;
		$counter_checked = FALSE;
		$span_repl_counter = 0;
		for($v = 0; $v < count($content_without_null_line); $v++) {
			$con = $content_without_null_line[$v];
			if(strchr($con, "<w:tc id=") == TRUE) {
				$column_counter++;
				$counter_checked = FALSE;
			}

			if(trim($con) == "</w:tc>") {
				$is_removable = FALSE;
				// $con = "";
			}

			if($is_removable == TRUE) {
				$con = "";
			}

			for($q = 0; $q < count($removable_cols); $q++) {
				if(strchr($con, "id=\"".$removable_cols[$q]."\"") == TRUE) {
					$is_removable = TRUE;
					$con = "";
				}
			}

			if(isset($table_id[$table_counter]) && strchr($con, "<w:tbl id=\"".$table_id[$table_counter]."\">") == TRUE) {
				$is_table_found = TRUE;
			}
			
			if($is_table_found == TRUE && strchr($con, "</w:tbl>") == TRUE) {
				$is_table_found = FALSE;
			}

			if($is_table_found == TRUE) {
				if(strchr($con, "<w:p id=") == TRUE || strchr($con, "</w:p>") == TRUE) {
					$con = str_replace("</w:p>","/-r/",$con);
					$con = str_replace("<w:p id","/r/",$con);
				}
			}

			$con = str_replace("</w:tbl>","/-table/",$con);
			$con = str_replace("</w:tr>","/-row/",$con);
			$con = str_replace("</w:tc>","/-col/",$con);
			for($m = 0; $m < count($table_id); $m++) {
				$con = str_replace("<w:tbl id=\"".$table_id[$m]."\">", "/table-/ style=\"border-collapse: collapse;\"border=\"1\" id=\"".$table_id[$m]."\"/link-/",$con);
			}
			for($n = 0; $n < count($table_column); $n++) {
				$con = str_replace("<w:tc id=\"".$table_column[$n]."\">", "/col-/ id=\"".$table_column[$n]."\"/link-/",$con);
			}
			for($o = 0; $o < count($table_row); $o++) {
				$con = str_replace("<w:tr id=\"".$table_row[$o]."\">", "/row-/ id=\"".$table_row[$o]."\"/link-/",$con);
			}

			for($a = 0; $a < count($row_spans); $a++) {
				$str_to_find = "id=\"".substr($row_spans[$a], 1, strlen($row_spans[$a]))."\"";
				if(strchr($con, $str_to_find)) {
					$span = $final_row_span[$a];
					$con = str_replace($str_to_find, $str_to_find." rowspan=\"".$final_row_span[$span][0]."\"", $con);
				}
			}

			// for($p = 0; $p < count($row_spans); $p++) {
			// 	if($row_spans[$p]) {
			// 		$str_to_find = "id=\"".substr($row_spans[$p], 1, strlen($row_spans[$p]))."\"";
			// 		$con = str_replace($str_to_find, $str_to_find." rowspan=\"".$row_spans[$p][0]."\"", $con);
			// 	}
			// }

			for($r = 0; $r < count($col_spans); $r++) {
				if($col_spans[$r]) {
					$str_to_find = "id=\"d".substr($col_spans[$r], 1, strlen($col_spans[$r]))."\"";
					$con = str_replace($str_to_find, $str_to_find." colspan=\"".$col_spans[$r][0]."\"", $con);
				}
			}

			$content_without_null_line[$v] = $con;
		}

		//the below code add a header tag to the elements
		$head_counter = 0;
		$head_start = FALSE;

		for($f = 0; $f < count($content_without_null_line); $f++) {
			$con = $content_without_null_line[$f];
			if(isset($header_id[$head_counter]) && (strchr($con, "<w:p id=\"".$header_id[$head_counter]."\">") || trim($con) == "</w:p>")) {
				if(trim($con) == "</w:p>") {
					if($head_start == TRUE) {
						$con = str_replace("</w:p>","</w:p>\n/-".$header_id_type[$head_counter]."/",$con);
						$head_counter++;
						$head_start = FALSE;
					}
				} else {
					$head_start = TRUE;
					$con = str_replace("<w:p id=\"".$header_id[$head_counter]."\">","/".$header_id_type[$head_counter]."/<w:p id=\"".$header_id[$head_counter]."\">",$con);
				}
			}
			$content_without_null_line[$f] = $con;
		}


		$html_doc = NULL;
		$last_sel = NULL;
		$list_counter = 0;
		$cur_list_id = 0;
		$list_opened = FALSE;
		$list_txt = array("ul","ol");
		
		if(isset($list_id_type[0])) {
			$cur_list_id = $list_id_type[0];
		}

		//the below code structures the list types
		foreach($content_without_null_line as $con) {
			if(isset($list_id[$list_counter]) && (strchr($con, "<w:p id=\"".$list_id[$list_counter]."\">") || trim($con) == "</w:p>")) {
				if(trim($con) == "</w:p>") {
					if($list_opened == TRUE) {
						if(count($list_id_type) - 1 == $list_counter) {
							$val = $list_id_type[$list_counter] - 1;
							$con = str_replace("</w:p>","/-li//-".$list_txt[$numberings[$val]]."/", $con);
							// echo "\tlast - ".$list_counter." - closing\n</".$list_txt[$numberings[$val]]."\n";
							$list_opened = FALSE;
						} else {
							$con = str_replace("</w:p>","/li/", $con);
							// echo "\tnull - ".$list_counter." - closing\n";
						}
						$list_opened = FALSE;
						$list_counter++;
					}
				} else {
					$list_opened = TRUE;
					if($list_counter == 0) {
						$val = $list_id_type[$list_counter] - 1;
						$last_sel = $list_txt[$numberings[$val]];
						$con = str_replace("<w:p id=\"".$list_id[$list_counter]."\">",
								"/".$last_sel."//li-/ id=\"".$list_id[$list_counter]."\"/link-/", $con);
						// echo "<".$list_txt[$numberings[$val]]."\n\tfirst- ".$list_counter." - opening\n";
					} else {
						if($cur_list_id != (int)$list_id_type[$list_counter]) {
							$cur_list_id = (int)$list_id_type[$list_counter];
							// echo "</".$last_sel."\n";
							$new_str = "/-".$last_sel."/";
							$val = $list_id_type[$list_counter] - 1;
							$last_sel = $list_txt[$numberings[$val]];

							$new_str .= "/".$last_sel."/";
							$con =  str_replace("<w:p id=\"".$list_id[$list_counter]."\">",
									$new_str."/li-/ id=\"".$list_id[$list_counter]."\"/link-/", $con);
							// echo "<".$last_sel."\n\tnull - ".$list_counter." - opening\n";
						} else {
							$con =  str_replace("<w:p id=\"".$list_id[$list_counter]."\">",
									"/li-/ id=\"".$list_id[$list_counter]."\"/link-/", $con);
							// echo "\tinline value - first also null - ".$list_counter." - opening\n";
						}
					}
				}
				$html_doc .= "\n".$con;
			} else if(strchr($con,"wp:position") || strchr($con,"wp:posOffset") || strchr($con,"wp:align")) {

			} else if(trim($con) != "") {
				$html_doc .= "\n".$con;
			}
		}

		foreach($styles as $sty) {
			$dups_checker = explode("\n",$sty);
			$new_str = "";
			for($e = 1; $e < count($dups_checker); $e++) {
				$new_str = str_replace("text-decoration-style: single","text-decoration: underline", $new_str);
				$new_str = str_replace("vertical-align: subscript","vertical-align: sub", $new_str);
				$new_str = str_replace("vertical-align: superscript","vertical-align: super", $new_str);
				if(trim($new_str) == "") {
					$new_str = "\n".$dups_checker[$e];
				} else if(trim($new_str) != "" && $dups_checker[$e] != "") {
					if(strchr($new_str, $dups_checker[$e]) == FALSE) {
						$new_str .= "\n".$dups_checker[$e];
					}
				}
			}

			if(trim($new_str) != "") {
				$html_doc .= "\n".$new_str;
			}
		}


		$html_doc_out_space = NULL;
		foreach(explode("\n",$this->tagReplacer($html_doc)) as $doc) {
			if(trim($doc) != "") {
				if(trim($doc) == "</body>" || trim($doc) == "<body>") {
				} else {
					$html_doc_out_space .= "\n\t".$doc;
				}
			}
		}

		
		$html_doc_out_space .= "\n\t\t<script>\n".
		"\t\t\t\$(\"li\").each(function() {\n".
		"\t\t\t\tif(!\$.trim(\$(this).html())) {\n".
		"\t\t\t\t\t\$(this).hide();\n".
		"\t\t\t\t}\n".
		"\t\t\t});\n".
		"\t\t\t\$(\"span\").each(function() {\n".
		"\t\t\t\tif(!\$.trim(\$(this).html())) {\n".
		"\t\t\t\t\t\$(this).hide();\n".
		"\t\t\t\t}\n".
		"\t\t\t});\n".
		"\t\t</script>\n";
		echo $html_doc_out_space;
	}

	public function tagReplacer($html_doc) {

		$html_doc = str_replace("<w:body>","/body/",$html_doc);
		$html_doc = str_replace("</w:body>","/-body/",$html_doc);
		$html_doc = str_replace("<w:p id","/p/",$html_doc);
		$html_doc = str_replace("</w:p>","/-p/",$html_doc);
		$html_doc = str_replace("<w:r id","/r/",$html_doc);
		$html_doc = str_replace("</w:r>","/-r/",$html_doc);
		$html_doc = str_replace("<w:t id","/t/",$html_doc);
		$html_doc = str_replace("</w:t>","/-t/",$html_doc);
		$html_doc = str_replace("<style>","/style/",$html_doc);
		$html_doc = str_replace("</style>","/-style/",$html_doc);
		$html_doc = str_replace("</w:hyperlink>","/-link/",$html_doc);
		$html_doc = str_replace("<w:hyperlink>","/link/",$html_doc);
		$html_doc = str_replace("<w:drawing","/img/",$html_doc);
		$html_doc = strip_tags($html_doc);
		
		$html_doc = str_replace("/body/","<body>",$html_doc);
		$html_doc = str_replace("/-body/","</body>",$html_doc);
		$html_doc = str_replace("/p/","<span id",$html_doc);
		$html_doc = str_replace("/-p/","</span><br>",$html_doc);
		$html_doc = str_replace("/r/","<span id",$html_doc);
		$html_doc = str_replace("/-r/","</span>",$html_doc);
		$html_doc = str_replace("/t/","<span id",$html_doc);
		$html_doc = str_replace("/-t/","</span>",$html_doc);
		$html_doc = str_replace("/style/","<style>",$html_doc);
		$html_doc = str_replace("/-style/","</style>",$html_doc);
		$html_doc = str_replace("/link/","<a",$html_doc);
		$html_doc = str_replace("/link-/",">",$html_doc);
		$html_doc = str_replace("/-link/","</a>",$html_doc);
		$html_doc = str_replace("/img/","<img",$html_doc);
		$html_doc = str_replace("/li/","<li>",$html_doc);
		$html_doc = str_replace("/-li/","</li>",$html_doc);
		$html_doc = str_replace("/li-/","<li ",$html_doc);
		$html_doc = str_replace("/ul/","<ul>",$html_doc);
		$html_doc = str_replace("/-ul/","</ul>",$html_doc);
		$html_doc = str_replace("/ol/","<ol>",$html_doc);
		$html_doc = str_replace("/-ol/","</ol>",$html_doc);
		$html_doc = str_replace("/h1/","<h1>",$html_doc);
		$html_doc = str_replace("/-h1/","</h1>",$html_doc);
		$html_doc = str_replace("/h2/","<h2>",$html_doc);
		$html_doc = str_replace("/-h2/","</h2>",$html_doc);
		$html_doc = str_replace("/h3/","<h3>",$html_doc);
		$html_doc = str_replace("/-h3/","</h3>",$html_doc);
		$html_doc = str_replace("/h4/","<h4>",$html_doc);
		$html_doc = str_replace("/-h4/","</h4>",$html_doc);
		$html_doc = str_replace("/h5/","<h5>",$html_doc);
		$html_doc = str_replace("/-h5/","</h5>",$html_doc);
		$html_doc = str_replace("/h6/","<h6>",$html_doc);
		$html_doc = str_replace("/-h6/","</h6>",$html_doc);
		$html_doc = str_replace("/-table/","</table>",$html_doc);
		$html_doc = str_replace("/-row/","</tr>",$html_doc);
		$html_doc = str_replace("/-col/","</td>",$html_doc);
		$html_doc = str_replace("/table-/","<table",$html_doc);
		$html_doc = str_replace("/row-/","<tr",$html_doc);
		$html_doc = str_replace("/col-/","<td",$html_doc);
		return $html_doc;
	}

	public function convertToText() {
		if(isset($this->docxFileName) && !file_exists($this->docxFileName)) {
			return "File Not exists";
		}
		
		$fileArray = pathinfo($this->docxFileName);
		$file_ext  = $fileArray['extension'];
		if($file_ext == "docx") {
			if($file_ext == "docx") {
				return $this->read_docx();
			}
		} else {
			return "We Only accept docx files at this time";
		}
	}
}
?>
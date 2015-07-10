<?php

/**
 * Description of xls_export
 *
 * @author Daniel
 */
class xls_export {
    private $column;
    private $row;
    private $cell;
    private $xlcontent;
    private $colmax;
    private $rowmax;
    private $msxml;
    private $cellstyle;
    private $cellcolspan;
    private $cellrowspan;
    private $date_time;
    private $author;
    private $pagebreak;
    private $worksheet;

    public function __construct() {
        $this->pagebreak = 0;
        $this->date_time = date("d.m.Y - H:i");
        $this->msxml =  "<html xmlns:v='urn:schemas-microsoft-com:vml'". 
                        "xmlns:o='urn:schemas-microsoft-com:office:office'".
                        "xmlns:x='urn:schemas-microsoft-com:office:excel'".
                        "xmlns='urn:schemas-microsoft-com:office:spreadsheet'>".
                        "xmlns:ss='urn:schemas-microsoft-com:office:spreadsheet'>".
                            "<head>".
                                "<meta http-equiv=Content-Type content='text/html; charset=UTF-8'>".
                                "<meta name=ProgId content=Excel.Sheet>".
                                "<meta name=Generator content='Microsoft Excel 11'>".
                                "<meta name=Originator content='Microsoft Excel 11'>".
                                "<xml>".
                                "<o:DocumentProperties>".
                                    "<o:Author>N.N</o:Author>".
                                    "<o:LastAuthor>N.N</o:LastAuthor>".
                                    "<o:Company>-</o:Company>".
                                    "<o:Created>".$this->date_time."</o:Created>".
                                    "<o:Version>11.9999</o:Version>".
                                "</o:DocumentProperties>".
                                "<x:ExcelWorkbook>".
                                    "<x:WindowHeight>13225</x:WindowHeight>".
                                    "<x:<WindowWidth>19382</x:WindowWidth>".
                                "</x:ExcelWorkbook>";
    }
    
    /**
     *
     * @param type $cell - The Cell in which the PageBreak should occur
     */
    public function setPageBreak($cell){
        $this->pagebreak = $cell;
        $this->msxml.=  "<x:RowBreak>".
                            "<x:Row>".$this->pagebreak."</x:Row>".
                        "</x:RowBreak>";    
    }
    
    /**
     * 
     * @param type $strcell - coordinates of the cell in X:Y
     * @param type $content - content for that cell
     */
    public function add_cell($strcell, $content){
        if(strpos($strcell,":") != false){
            list($this->column,$this->row) = explode(":",$strcell);
            $this->cell[$this->column][$this->row]=$content;
            if($this->column > $this->colmax){
                $this->colmax = $this->column;
            }
            if($this->row > $this->rowmax){
                $this->rowmax = $this->row;
            }
        }        
    }
    
    /**
     * 
     * @param type $worksheet - The name of the worksheet to use
     */
    public function setWorkSheet($worksheet,$pagebreak=0){
        $this->worksheet = $worksheet;
        $this->pagebreak = $pagebreak;
        $this->msxml.=  "<x:ExcelWorksheet ss:Name=".$this->worksheet.">".
                            "<x:PageBreaks>".
                            "<x:RowBreaks>";
        $this->setPageBreak($this->pagebreak);
    }
}

module OoxmlParser
  # Class for `tableStyleInfo` data
  # https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.tablestyleinfo.aspx
  class TableStyleInfo < OOXMLDocumentObject
    # @return [String] A string representing the name of the table style to use with this table.
    # If the style name does not correspond to the name of a table style then the spreadsheet
    # application should use default style.
    # The possible values for this attribute are defined by the ST_Xstring simple type (22.9.2.19).
    attr_accessor :name

    # @return [True, False] A Boolean indicating whether the first column in the table should have the style applied.
    attr_accessor :show_first_column
    # @return [True, False] A Boolean indicating whether column stripe formatting is applied.
    attr_accessor :show_column_stripes
    # @return [True, False] A Boolean indicating whether the last column in the table should have the style applied.
    attr_accessor :show_last_column
    # @return [True, False] A Boolean indicating whether row stripe formatting is applied.
    attr_accessor :show_row_stripes

    # Parse TableStyleInfo data
    # @param [Nokogiri::XML:Element] node with TableStyleInfo data
    # @return [TableStyleInfo] value of TableStyleInfo data
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'name'
          @name = value.value.to_s
        when 'showColumnStripes'
          @show_column_stripes = attribute_enabled?(value)
        when 'showFirstColumn'
          @show_first_column = attribute_enabled?(value)
        when 'showLastColumn'
          @show_last_column = attribute_enabled?(value)
        when 'showRowStripes'
          @show_row_stripes = attribute_enabled?(value)
        end
      end
      self
    end
  end
end

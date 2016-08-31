module OoxmlParser
  # Class for parsing `w:tblLayout` object
  # Whether a table uses a fixed width or autofit method for laying out the table
  # contents is specified with the <w:tblLayout> element within the <w:tblPr> element.
  # If <w:tblLayout> is omitted, autofit is assumed.
  class TableLayout < OOXMLDocumentObject
    # @return [Symbol] Specifies the method of laying out the contents of the table
    attr_accessor :type

    # Parse TableLayout
    # @param [Nokogiri::XML:Node] node with TableLayout
    # @return [TableLayout] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'type'
          @type = value.value.to_sym
        end
      end
      self
    end
  end
end

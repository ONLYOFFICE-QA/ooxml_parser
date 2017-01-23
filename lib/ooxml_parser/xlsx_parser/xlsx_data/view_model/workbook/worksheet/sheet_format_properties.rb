module OoxmlParser
  # Class for storing SheetFormatProperties
  class SheetFormatProperties < OOXMLDocumentObject
    attr_accessor :default_column_width, :default_row_height

    # Parse SheetFormatProperties
    # @param node [Nokogiri::XML::Element] node to parse
    # @return [SheetFormatProperties] value of SheetFormatProperties
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'defaultColWidth'
          @default_column_width = value.value.to_f
        when 'defaultRowHeight'
          @default_row_height = value.value.to_f
        end
      end
      self
    end
  end
end

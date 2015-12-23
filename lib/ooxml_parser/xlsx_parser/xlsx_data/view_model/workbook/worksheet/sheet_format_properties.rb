# Class for storing SheetFormatProperties
module OoxmlParser
  class SheetFormatProperties
    attr_accessor :default_column_width, :default_row_height

    def initialize(column_width = nil, row_height = nil)
      @default_column_width = column_width
      @default_row_height = row_height
    end

    # Parse SheetFormatProperties
    # @param node [Nokogiri::XML::Element] node to parse
    # @return [SheetFormatProperties] value of SheetFormatProperties
    def self.parse(node)
      format_properties = SheetFormatProperties.new
      format_properties.default_column_width = node.attribute('defaultColWidth').value.to_f
      format_properties.default_row_height = node.attribute('defaultRowHeight').value.to_f
      format_properties
    end
  end
end

# frozen_string_literal: true

module OoxmlParser
  # Class for parsing `w:gridCol` object
  class GridColumn < OOXMLDocumentObject
    # @return [Float] width of column
    attr_accessor :width

    # Parse GridColumn
    # @param [Nokogiri::XML:Node] node with GridColumn
    # @return [GridColumn] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'w'
          @width = OoxmlSize.new(value.value.to_f, :emu)
        end
      end
      self
    end
  end
end

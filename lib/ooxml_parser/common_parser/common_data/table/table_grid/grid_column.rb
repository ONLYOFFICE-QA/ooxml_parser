module OoxmlParser
  # Class for parsing `w:gridCol` object
  class GridColumn < OOXMLDocumentObject
    # @return [Float] width of column
    attr_accessor :width

    # Parse GridColumn
    # @param [Nokogiri::XML:Node] node with GridColumn
    # @return [GridColumn] result of parsing
    def self.parse(node)
      grid_column = GridColumn.new

      node.attributes.each do |key, value|
        case key
        when 'w'
          grid_column.width = (value.value.to_f / OoxmlParser.configuration.units_delimiter).round(2)
        end
      end
      grid_column
    end
  end
end

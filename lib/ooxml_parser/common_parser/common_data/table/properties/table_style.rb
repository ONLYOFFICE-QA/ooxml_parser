require_relative 'table_element'
# Style of Table
module OoxmlParser
  class TableStyle < OOXMLDocumentObject
    attr_accessor :id, :name, :whole_table, :banding_1_horizontal, :banding_2_horizontal, :banding_1_vertical, :banding_2_vertical,
                  :last_column, :first_column, :last_row, :first_row, :southeast_cell, :southwest_cell, :northeast_cell,
                  :northwest_cell
    # @return [String] value of table style
    attr_reader :value

    # Parse TableStyle object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [TableStyle] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'styleId'
          @id = value.value.to_s
        when 'styleName'
          @name = value.value.to_s
        when 'val'
          @value = value.value.to_s
        end
      end

      node.xpath('*').each do |style_node_child|
        case style_node_child.name
        when 'wholeTbl'
          @whole_table = TableElement.new(parent: self).parse(style_node_child)
        when 'band1H'
          @banding_1_horizontal = TableElement.new(parent: self).parse(style_node_child)
        when 'band2H', 'band2Horz'
          @banding_2_horizontal = TableElement.new(parent: self).parse(style_node_child)
        when 'band1V'
          @banding_1_vertical = TableElement.new(parent: self).parse(style_node_child)
        when 'band2V'
          @banding_2_vertical = TableElement.new(parent: self).parse(style_node_child)
        when 'lastCol'
          @last_column = TableElement.new(parent: self).parse(style_node_child)
        when 'firstCol'
          @first_column = TableElement.new(parent: self).parse(style_node_child)
        when 'lastRow'
          @last_row = TableElement.new(parent: self).parse(style_node_child)
        when 'firstRow'
          @first_row = TableElement.new(parent: self).parse(style_node_child)
        when 'seCell'
          @southeast_cell = TableElement.new(parent: self).parse(style_node_child)
        when 'swCell'
          @southwest_cell = TableElement.new(parent: self).parse(style_node_child)
        when 'neCell'
          @northeast_cell = TableElement.new(parent: self).parse(style_node_child)
        when 'nwCell'
          @northwest_cell = TableElement.new(parent: self).parse(style_node_child)
        end
      end
      self
    end
  end
end

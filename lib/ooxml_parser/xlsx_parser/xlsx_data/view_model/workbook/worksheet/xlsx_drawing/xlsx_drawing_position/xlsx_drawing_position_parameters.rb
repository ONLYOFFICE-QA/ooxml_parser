module OoxmlParser
  class XlsxDrawingPositionParameters
    attr_accessor :column, :column_offset, :row, :row_offset

    def self.parse(position_node)
      drawing_position_parameters = XlsxDrawingPositionParameters.new
      position_node.xpath('*').each do |position_node_child|
        case position_node_child.name
        when 'col'
          drawing_position_parameters.column = Coordinates.get_column_name(position_node_child.text.to_i + 1)
        when 'colOff'
          drawing_position_parameters.column_offset = (position_node_child.text.to_f / 360_000.0).round(3)
        when 'row'
          drawing_position_parameters.row = position_node_child.text.to_i + 1
        when 'rowOff'
          drawing_position_parameters.row_offset = (position_node_child.text.to_f / 360_000.0).round(3)
        end
      end
      drawing_position_parameters
    end
  end
end

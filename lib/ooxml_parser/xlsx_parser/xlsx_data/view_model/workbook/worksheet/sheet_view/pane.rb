# XLSX data of pane
module OoxmlParser
  class Pane
    attr_accessor :state, :top_left_cell, :x_split, :y_split

    def self.parse(pane_node)
      pane = Pane.new
      pane_node.attributes.each do |key, value|
        case key
        when 'state'
          pane.state = value.value.to_sym
        when 'topLeftCell'
          pane.top_left_cell = Coordinates.parse_coordinates_from_string(value.value)
        when 'xSplit'
          pane.x_split = value.value
        when 'ySplit'
          pane.y_split = value.value
        end
      end
      pane
    end
  end
end

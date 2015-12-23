# Legend of Chart
module OoxmlParser
  class ChartLegend
    attr_accessor :position, :overlay

    def initialize(position = :right, overlay = false)
      @position = position
      @overlay = overlay
    end

    # Return combined data from @position and @overlay
    # If there is no overlay - return :right f.e.
    # If there is overlay - return :right_overlay
    # @return [Symbol] overlay and position type
    def position_with_overlay
      return "#{@position}_overlay".to_sym if overlay
      @position
    end

    def self.parse(legend_node)
      legend = ChartLegend.new
      legend_node.xpath('*').each do |legend_node_child|
        case legend_node_child.name
        when 'legendPos'
          legend.position = Alignment.parse(legend_node_child.attribute('val'))
        when 'overlay'
          legend.overlay = true if legend_node_child.attribute('val').value.to_s == '1'
        end
      end
      legend
    end
  end
end

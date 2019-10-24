# frozen_string_literal: true

module OoxmlParser
  # Legend of Chart `legend` tag
  class ChartLegend < OOXMLDocumentObject
    attr_accessor :position, :overlay

    def initialize(position = :right,
                   overlay = false,
                   parent: nil)
      @position = position
      @overlay = overlay
      @parent = parent
    end

    # Return combined data from @position and @overlay
    # If there is no overlay - return :right f.e.
    # If there is overlay - return :right_overlay
    # @return [Symbol] overlay and position type
    def position_with_overlay
      return "#{@position}_overlay".to_sym if overlay

      @position
    end

    # Parse ChartLegend object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [ChartLegend] result of parsing
    def parse(node)
      node.xpath('*').each do |child_node|
        case child_node.name
        when 'legendPos'
          @position = value_to_symbol(child_node.attribute('val'))
        when 'overlay'
          @overlay = true if child_node.attribute('val').value.to_s == '1'
        end
      end
      self
    end
  end
end

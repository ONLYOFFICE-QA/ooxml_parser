# frozen_string_literal: true

module OoxmlParser
  # Class for `selection` data
  class Selection < OOXMLDocumentObject
    # @return [Coordinates] Reference to the active cell
    attr_reader :active_cell
    # @return [Integer] Id of active cell
    attr_reader :active_cell_id
    # @return [String] Selected range
    attr_reader :reference_sequence

    # Parse Selection object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [Selection] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'activeCell'
          @active_cell = Coordinates.new.parse_string(value.value)
        when 'activeCellId'
          @active_cell_id = value.value.to_i
        when 'sqref'
          @reference_sequence = value.value.to_s
        end
      end
      self
    end
  end
end

# frozen_string_literal: true

module OoxmlParser
  # Class to describe spacing valued child
  class SpacingValuedChild < OOXMLDocumentObject
    # @return [ValuedChild] spacing percent
    attr_reader :spacing_percent
    # @return [ValuedChild] spacing point
    attr_reader :spacing_points

    # Parse LineSpacing object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [LineSpacing] result of parsing
    def parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'spcPct'
          @spacing_percent = ValuedChild.new(:float, parent: self).parse(node_child)
        when 'spcPts'
          @spacing_points = ValuedChild.new(:float, parent: self).parse(node_child)
        end
      end
      self
    end

    # Convert to OoxmlSize
    # @return [OoxmlSize]
    def to_ooxml_size
      return OoxmlSize.new(@spacing_percent.value, :one_1000th_percent) if @spacing_percent
      return OoxmlSize.new(@spacing_points.value, :spacing_point) if @spacing_points

      raise 'Unknown spacing child type'
    end

    # @return [Symbol] rule used to determine line spacing
    def rule
      return :multiple if @spacing_percent

      :exact
    end
  end
end

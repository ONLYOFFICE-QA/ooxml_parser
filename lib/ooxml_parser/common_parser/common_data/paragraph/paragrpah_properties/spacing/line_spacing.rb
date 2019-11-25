# frozen_string_literal: true

module OoxmlParser
  # Class to describe spacing
  class LineSpacing < OOXMLDocumentObject
    # @return [ValuedChild] spacing percent
    attr_reader :spacing_percent
    # @return [ValuedChild] spacing point
    attr_reader :@pacing_points

    # Parse LineSpacing object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [LineSpacing] result of parsing
    def parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'spcPct'
          @spacing_percent = ValuedChild.new(:integer, parent: self).parse(node_child)
        when 'spcPts'
          @spacing_points = ValuedChild.new(:integer, parent: self).parse(node_child)
        end
      end
      self
    end

    # @return [Symbol] rule used to determine line spacing
    def rule
      return :multiple if @spacing_percent

      :exact
    end
  end
end

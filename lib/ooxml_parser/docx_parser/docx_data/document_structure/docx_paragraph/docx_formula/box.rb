# frozen_string_literal: true

module OoxmlParser
  # Class for `box`, `borderBox` data
  class Box < OOXMLDocumentObject
    attr_accessor :borders, :element

    def initialize(parent: nil)
      @borders = false
      @parent = parent
    end

    # Parse Box object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [Box] result of parsing
    def parse(node)
      @borders = true if node.name == 'borderBox'
      @element = MathRun.new(parent: self).parse(node)
      self
    end
  end
end

# frozen_string_literal: true

module OoxmlParser
  # Class for parsing `w:pBdr` element
  class ParagraphBorders < OOXMLDocumentObject
    # @return [BordersProperties] bottom properties
    attr_accessor :bottom
    # @return [BordersProperties] left properties
    attr_accessor :left
    # @return [BordersProperties] top properties
    attr_accessor :top
    # @return [BordersProperties] right properties
    attr_accessor :right
    # @return [BordersProperties] between properties
    attr_accessor :between
    # @return [BordersProperties] bar properties
    attr_accessor :bar

    # @return [Symbol] type of border in visual editor
    def border_visual_type
      result = []
      result << :left if @left.val == :single
      result << :right if @right.val == :single
      result << :top if @top.val == :single
      result << :bottom if @bottom.val == :single
      result << :inner if @between.val == :single
      return :none if result == []
      return :all if result == %i[left right top bottom inner]
      return :outer if result == %i[left right top bottom]

      result.first if result.size == 1
    end

    # Parse Paragraph Borders data
    # @param [Nokogiri::XML:Element] node with Paragraph Borders data
    # @return [ParagraphBorders] value of Paragraph Borders data
    def parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'bottom'
          @bottom = BordersProperties.new(parent: self).parse(node_child)
        when 'left'
          @left = BordersProperties.new(parent: self).parse(node_child)
        when 'top'
          @top = BordersProperties.new(parent: self).parse(node_child)
        when 'right'
          @right = BordersProperties.new(parent: self).parse(node_child)
        when 'between'
          @between = BordersProperties.new(parent: self).parse(node_child)
        when 'bar'
          @bar = BordersProperties.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end

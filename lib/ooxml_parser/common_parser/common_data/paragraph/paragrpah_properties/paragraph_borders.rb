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

    def border_visual_type
      result = []
      result << :left if @left.val == :single
      result << :right if @right.val == :single
      result << :top if @top.val == :single
      result << :bottom if @bottom.val == :single
      result << :inner if @between.val == :single
      return :none if result == []
      return :all if result == [:left, :right, :top, :bottom, :inner]
      return :outer if result == [:left, :right, :top, :bottom]
      result.first if result.size == 1
    end

    # Parse Paragraph Borders data
    # @param [Nokogiri::XML:Element] node with Paragraph Borders data
    # @return [ParagraphBorders] value of Paragraph Borders data
    def self.parse(node)
      paragraph_borders = ParagraphBorders.new
      node.xpath('*').each do |border_pr_child|
        case border_pr_child.name
        when 'bottom'
          paragraph_borders.bottom = BordersProperties.new(parent: paragraph_borders).parse(border_pr_child)
        when 'left'
          paragraph_borders.left = BordersProperties.new(parent: paragraph_borders).parse(border_pr_child)
        when 'top'
          paragraph_borders.top = BordersProperties.new(parent: paragraph_borders).parse(border_pr_child)
        when 'right'
          paragraph_borders.right = BordersProperties.new(parent: paragraph_borders).parse(border_pr_child)
        when 'between'
          paragraph_borders.between = BordersProperties.new(parent: paragraph_borders).parse(border_pr_child)
        when 'bar'
          paragraph_borders.bar = BordersProperties.new(parent: paragraph_borders).parse(border_pr_child)
        end
      end
      paragraph_borders
    end
  end
end

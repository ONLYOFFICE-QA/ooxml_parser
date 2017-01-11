module OoxmlParser
  # Class for parsing `pgMar` tags
  class PageMargins < OOXMLDocumentObject
    attr_accessor :top, :bottom, :left, :right, :footer, :gutter, :header

    def initialize(top: nil,
                   bottom: nil,
                   left: nil,
                   right: nil,
                   header: nil,
                   footer: nil,
                   gutter: nil,
                   parent: nil)
      @top = top
      @bottom = bottom
      @left = left
      @right = right
      @header = header
      @footer = footer
      @gutter = gutter
      @parent = parent
    end

    # Parse BordersProperties
    # @param [Nokogiri::XML:Element] node with PageMargins
    # @return [PageMargins] value of PageMargins
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'top'
          @top = OoxmlSize.new(value.value.to_f)
        when 'left'
          @left = OoxmlSize.new(value.value.to_f)
        when 'right'
          @right = OoxmlSize.new(value.value.to_f)
        when 'bottom'
          @bottom = OoxmlSize.new(value.value.to_f)
        when 'header'
          @header = OoxmlSize.new(value.value.to_f)
        when 'footer'
          @footer = OoxmlSize.new(value.value.to_f)
        when 'gutter'
          @gutter = OoxmlSize.new(value.value.to_f)
        end
      end
      self
    end
  end
end

module OoxmlParser
  class PageMargins < OOXMLDocumentObject
    attr_accessor :top, :bottom, :left, :right, :footer, :gutter, :header

    def initialize(top: nil, bottom: nil, left: nil, right: nil, header: nil, footer: nil, gutter: nil)
      @top = top
      @bottom = bottom
      @left = left
      @right = right
      @header = header
      @footer = footer
      @gutter = gutter
    end

    # Parse BordersProperties
    # @param [Nokogiri::XML:Element] node with PageMargins
    # @return [PageMargins] value of PageMargins
    def self.parse(node)
      margins = PageMargins.new
      node.attributes.each do |key, value|
        case key
        when 'top'
          margins.top = OoxmlSize.new(value.value.to_f)
        when 'left'
          margins.left = OoxmlSize.new(value.value.to_f)
        when 'right'
          margins.right = OoxmlSize.new(value.value.to_f)
        when 'bottom'
          margins.bottom = OoxmlSize.new(value.value.to_f)
        when 'header'
          margins.header = OoxmlSize.new(value.value.to_f)
        when 'footer'
          margins.footer = OoxmlSize.new(value.value.to_f)
        when 'gutter'
          margins.gutter = OoxmlSize.new(value.value.to_f)
        end
      end
      margins
    end
  end
end

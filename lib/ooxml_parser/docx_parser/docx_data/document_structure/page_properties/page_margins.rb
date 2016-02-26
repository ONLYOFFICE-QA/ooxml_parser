module OoxmlParser
  class PageMargins
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

    def ==(other)
      other.class == PageMargins &&
        @top == other.top &&
        @bottom == other.bottom &&
        @left == other.left &&
        @right == other.right &&
        @footer == other.footer &&
        @header == other.header &&
        @gutter == other.gutter
    end

    # Parse BordersProperties
    # @param [Nokogiri::XML:Element] node with PageMargins
    # @return [PageMargins] value of PageMargins
    def self.parse(node)
      margins = PageMargins.new
      # TODO: implement and understand, why 566.929, but not `unit_delimiter`
      margins.top = (node.attribute('top').value.to_f / 566.929).round(2)
      margins.left = (node.attribute('left').value.to_f / 566.929).round(2)
      margins.right = (node.attribute('right').value.to_f / 566.929).round(2)
      margins.bottom = (node.attribute('bottom').value.to_f / 566.929).round(2)
      margins.header = (node.attribute('header').value.to_f / 566.929).round(2) unless node.attribute('header').nil?
      margins.footer = (node.attribute('footer').value.to_f / 566.929).round(2) unless node.attribute('footer').nil?
      margins.gutter = (node.attribute('gutter').value.to_f / 566.929).round(2)
      margins
    end
  end
end

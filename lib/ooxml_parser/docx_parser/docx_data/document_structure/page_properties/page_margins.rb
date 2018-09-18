module OoxmlParser
  # Class for parsing `pgMar` tags
  class PageMargins < OOXMLDocumentObject
    attr_accessor :top, :bottom, :left, :right, :footer, :gutter, :header

    def initialize(params)
      @top = params[:top]
      @bottom = params[:bottom]
      @left = params[:left]
      @right = params[:right]
      @header = params[:header]
      @footer = params[:footer]
      @gutter = params[:gutter]
      @parent = params[:parent]
    end

    # Parse BordersProperties
    # @param [Nokogiri::XML:Element] node with PageMargins
    # @return [PageMargins] value of PageMargins
    def parse(node, unit = :dxa)
      node.attributes.each do |key, value|
        case key
        when 'top'
          @top = OoxmlSize.new(value.value.to_f, unit)
        when 'left'
          @left = OoxmlSize.new(value.value.to_f, unit)
        when 'right'
          @right = OoxmlSize.new(value.value.to_f, unit)
        when 'bottom'
          @bottom = OoxmlSize.new(value.value.to_f, unit)
        when 'header'
          @header = OoxmlSize.new(value.value.to_f, unit)
        when 'footer'
          @footer = OoxmlSize.new(value.value.to_f, unit)
        when 'gutter'
          @gutter = OoxmlSize.new(value.value.to_f, unit)
        end
      end
      self
    end
  end
end

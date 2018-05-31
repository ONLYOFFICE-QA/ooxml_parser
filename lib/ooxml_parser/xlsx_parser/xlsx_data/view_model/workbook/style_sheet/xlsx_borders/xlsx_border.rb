module OoxmlParser
  # Class for parsing `border`
  class XlsxBorder < OOXMLDocumentObject
    # @return [Border] value of border
    attr_reader :left
    # @return [Border] value of border
    attr_reader :right
    # @return [Border] value of border
    attr_reader :top
    # @return [Border] value of border
    attr_reader :bottom

    # Parse XlsxBorder data
    # @param [Nokogiri::XML:Element] node with XlsxBorder data
    # @return [XlsxBorder] value of XlsxBorder data
    def parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'bottom'
          @bottom = Border.new(parent: self).parse(node_child)
        when 'top'
          @top = Border.new(parent: self).parse(node_child)
        when 'right'
          @right = Border.new(parent: self).parse(node_child)
        when 'left'
          @left = Border.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end

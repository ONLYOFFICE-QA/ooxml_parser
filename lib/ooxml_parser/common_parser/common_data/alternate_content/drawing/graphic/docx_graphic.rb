require_relative 'shape/docx_shape'
require_relative 'group/docx_grouped_drawing'
require_relative 'picture/docx_picture'
module OoxmlParser
  # Class for parsing `graphic` tags
  class DocxGraphic < OOXMLDocumentObject
    attr_accessor :type, :data

    alias chart data

    # Parse DocxGraphic
    # @param [Nokogiri::XML:Node] node with NumberingProperties
    # @return [DocxGraphic] result of parsing
    def parse(node)
      node.xpath('a:graphicData/*', 'xmlns:a' => 'http://schemas.openxmlformats.org/drawingml/2006/main').each do |node_child|
        case node_child.name
        when 'wsp'
          @type = :shape
          @data = DocxShape.new(parent: self).parse(node_child)
        when 'pic'
          @type = :picture
          @data = DocxPicture.new(parent: self).parse(node_child)
        when 'chart'
          @type = :chart
          OOXMLDocumentObject.add_to_xmls_stack("#{OOXMLDocumentObject.root_subfolder}/#{OOXMLDocumentObject.get_link_from_rels(node_child.attribute('id').value)}")
          @data = Chart.parse
          OOXMLDocumentObject.xmls_stack.pop
        when 'wgp'
          @type = :group
          @data = DocxGroupedDrawing.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end

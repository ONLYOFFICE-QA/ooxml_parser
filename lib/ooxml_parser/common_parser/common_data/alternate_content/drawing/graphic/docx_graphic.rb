require_relative 'shape/docx_shape'
require_relative 'group/docx_grouped_drawing'
require_relative 'picture/docx_picture'
# Docx Graphic Data
module OoxmlParser
  class DocxGraphic < OOXMLDocumentObject
    attr_accessor :type, :data

    alias chart data

    def self.parse(graphic_node, parent: nil)
      graphic = DocxGraphic.new
      graphic.parent = parent
      graphic_node.xpath('a:graphicData/*', 'xmlns:a' => 'http://schemas.openxmlformats.org/drawingml/2006/main').each do |graphic_data_node_child|
        case graphic_data_node_child.name
        when 'wsp'
          graphic.type = :shape
          graphic.data = DocxShape.new(parent: graphic).parse(graphic_data_node_child)
        when 'pic'
          graphic.type = :picture
          graphic.data = DocxPicture.new(parent: graphic).parse(graphic_data_node_child)
        when 'chart'
          graphic.type = :chart
          OOXMLDocumentObject.add_to_xmls_stack("#{OOXMLDocumentObject.root_subfolder}/#{OOXMLDocumentObject.get_link_from_rels(graphic_data_node_child.attribute('id').value)}")
          graphic.data = Chart.parse
          OOXMLDocumentObject.xmls_stack.pop
        when 'wgp'
          graphic.type = :group
          graphic.data = DocxGroupedDrawing.parse(graphic_data_node_child)
        end
      end
      graphic
    end
  end
end

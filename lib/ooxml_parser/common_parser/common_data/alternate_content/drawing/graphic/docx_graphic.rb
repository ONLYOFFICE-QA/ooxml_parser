# frozen_string_literal: true

require_relative 'shape/docx_shape'
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
          @chart_reference = ChartReference.new(parent: self).parse(node_child)
          root_object.add_to_xmls_stack("#{root_object.root_subfolder}/#{root_object.get_link_from_rels(@chart_reference.id)}")
          @data = Chart.new(parent: self).parse
          root_object.xmls_stack.pop
        when 'wgp'
          @type = :group
          @data = ShapesGrouping.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end

module OoxmlParser
  # Class for parsing `graphicFrame`
  class GraphicFrame < OOXMLDocumentObject
    attr_accessor :properties, :transform, :graphic_data
    # @return [NonVisualShapeProperties] non visual properties
    attr_accessor :non_visual_properties

    def initialize(graphic_data = [])
      @graphic_data = graphic_data
    end

    # Parse GraphicFrame object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [GraphicFrame] result of parsing
    def parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'xfrm'
          @transform = DocxShapeSize.new(parent: self).parse(node_child)
        when 'graphic'
          graphic_data = []
          node_child.xpath('*').each do |node_child_child|
            case node_child_child.name
            when 'graphicData'
              node_child_child.xpath('*').each do |graphic_node_child|
                case graphic_node_child.name
                when 'tbl'
                  graphic_data << Table.new(parent: self).parse(graphic_node_child)
                when 'chart'
                  OOXMLDocumentObject.add_to_xmls_stack(OOXMLDocumentObject.get_link_from_rels(graphic_node_child.attribute('id').value))
                  graphic_data << Chart.parse
                  OOXMLDocumentObject.xmls_stack.pop
                end
              end
            end
          end
          @graphic_data = graphic_data
        when 'nvGraphicFramePr'
          @non_visual_properties = NonVisualShapeProperties.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end

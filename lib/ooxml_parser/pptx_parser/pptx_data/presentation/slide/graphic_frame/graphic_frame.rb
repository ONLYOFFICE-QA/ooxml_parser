module OoxmlParser
  class GraphicFrame < OOXMLDocumentObject
    attr_accessor :properties, :transform, :graphic_data

    def initialize(graphic_data = [])
      @graphic_data = graphic_data
    end

    def self.parse(graphic_frame_node)
      graphic_frame = GraphicFrame.new
      graphic_frame_node.xpath('*').each do |graphic_frame_node_child|
        case graphic_frame_node_child.name
        when 'xfrm'
          graphic_frame.transform = TransformEffect.parse(graphic_frame_node_child)
        when 'graphic'
          graphic_data = []
          graphic_frame_node_child.xpath('a:graphicData/*').each do |graphic_node_child|
            case graphic_node_child.name
            when 'tbl'
              graphic_data << Table.new(parent: graphic_frame).parse(graphic_node_child)
            when 'chart'
              OOXMLDocumentObject.add_to_xmls_stack(OOXMLDocumentObject.get_link_from_rels(graphic_node_child.attribute('id').value))
              graphic_data << Chart.parse
              OOXMLDocumentObject.xmls_stack.pop
            end
          end
          graphic_frame.graphic_data = graphic_data
        end
      end
      graphic_frame
    end
  end
end

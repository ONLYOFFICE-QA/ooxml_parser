require_relative 'docx_blip'
# Docx Picture Data
module OoxmlParser
  class DocxPicture
    attr_accessor :path_to_image, :properties, :nonvisual_properties, :chart

    alias image path_to_image
    alias shape_properties properties

    def self.parse(picture_node)
      picture = DocxPicture.new
      picture_node.xpath('*').each do |picture_node_child|
        case picture_node_child.name
        when 'nvPicPr'
        when 'blipFill'
          picture.path_to_image = DocxBlip.parse(picture_node_child)
        when 'spPr'
          picture.properties = DocxShapeProperties.parse(picture_node_child)
        end
      end
      picture
    end
  end

  Picture = DocxPicture
end

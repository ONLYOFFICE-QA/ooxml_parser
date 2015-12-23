# Fallback DOCX shape fill properties
module OoxmlParser
  class OldDocxShapeFill < OOXMLDocumentObject
    attr_accessor :path_to_image, :stretching_type, :opacity, :title

    def self.parse(fill_node)
      fill = OldDocxShapeFill.new
      fill_node.attributes.each do |key, value|
        case key
        when 'id'
          fill.path_to_image = OOXMLDocumentObject.copy_media_file(OOXMLDocumentObject.root_subfolder + get_link_from_rels(value.value))
        when 'type'
          case value.value
          when 'frame'
            fill.stretching_type = :stretch
          else
            fill.stretching_type = value.value.to_sym
          end
        when 'title'
          fill.title = value.value.to_s
        end
      end
      fill
    end
  end
end

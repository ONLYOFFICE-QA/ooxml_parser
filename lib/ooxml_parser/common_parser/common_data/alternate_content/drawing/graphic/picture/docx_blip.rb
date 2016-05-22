# DOCX Blip properties
module OoxmlParser
  class DocxBlip < OOXMLDocumentObject
    attr_accessor :path_to_media_file, :alpha_channel

    alias path path_to_media_file

    def to_str
      path_to_media_file
    end

    def self.parse(blip_fill_node)
      blip = DocxBlip.new
      blip_node = blip_fill_node.xpath('a:blip', 'xmlns:a' => 'http://schemas.openxmlformats.org/drawingml/2006/main').first
      return blip if blip_node.nil?
      path_to_media_file = OOXMLDocumentObject.get_link_from_rels(blip_node.attribute('embed').value)
      raise LoadError, "Cant find path to media file by id: #{blip_node.attribute('embed').value}" if path_to_media_file.empty?
      blip_node.xpath('*').each do |blip_node_child|
        case blip_node_child.name
        when 'alphaModFix'
          blip.alpha_channel = (blip_node_child.attribute('amt').value.to_f / 1_000.0).round(0).to_f
        end
      end
      blip.path_to_media_file = OOXMLDocumentObject.copy_media_file("#{OOXMLDocumentObject.root_subfolder}/#{path_to_media_file.gsub('..', '')}") unless path_to_media_file == 'NULL'
      blip
    end
  end
end

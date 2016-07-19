require_relative 'docx_blip/file_reference'
module OoxmlParser
  # Docx Blip properties
  class DocxBlip < OOXMLDocumentObject
    attr_accessor :path_to_media_file, :alpha_channel
    # @return [FileReference] image structure
    attr_accessor :file_reference

    alias path path_to_media_file

    def to_str
      path_to_media_file
    end

    def self.parse(blip_fill_node)
      blip = DocxBlip.new
      blip_node = blip_fill_node.xpath('a:blip', 'xmlns:a' => 'http://schemas.openxmlformats.org/drawingml/2006/main').first
      return blip if blip_node.nil?
      blip_node.xpath('*').each do |blip_node_child|
        case blip_node_child.name
        when 'alphaModFix'
          blip.alpha_channel = (blip_node_child.attribute('amt').value.to_f / 1_000.0).round(0).to_f
        end
      end
      blip.file_reference = FileReference.parse(blip_node)
      blip
    end
  end
end

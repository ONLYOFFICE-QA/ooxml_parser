require_relative 'docx_blip/file_reference'
module OoxmlParser
  # Class for parsing `blipFill`
  class DocxBlip < OOXMLDocumentObject
    attr_accessor :path_to_media_file, :alpha_channel
    # @return [FileReference] image structure
    attr_accessor :file_reference

    alias path path_to_media_file

    def to_str
      path_to_media_file
    end

    # Parse DocxBlip object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [DocxBlip] result of parsing
    def parse(node)
      blip_node = node.xpath('a:blip', 'xmlns:a' => 'http://schemas.openxmlformats.org/drawingml/2006/main').first
      return self if blip_node.nil?
      blip_node.xpath('*').each do |node_child|
        case node_child.name
        when 'alphaModFix'
          @alpha_channel = (node_child.attribute('amt').value.to_f / 1_000.0).round(0).to_f
        end
      end
      @file_reference = FileReference.new(parent: self).parse(blip_node)
      self
    end
  end
end

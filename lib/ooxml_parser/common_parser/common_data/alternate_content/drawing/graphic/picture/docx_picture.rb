require_relative 'docx_blip'
module OoxmlParser
  # Class for parsing `pic`
  class DocxPicture < OOXMLDocumentObject
    attr_accessor :path_to_image, :properties, :nonvisual_properties, :chart

    alias image path_to_image
    alias shape_properties properties

    # Parse DocxPicture object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [DocxPicture] result of parsing
    def parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'blipFill'
          @path_to_image = DocxBlip.new(parent: self).parse(node_child)
        when 'spPr'
          @properties = DocxShapeProperties.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end

  Picture = DocxPicture
end

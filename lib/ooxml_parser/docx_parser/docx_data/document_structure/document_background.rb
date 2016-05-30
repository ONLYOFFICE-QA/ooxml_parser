module OoxmlParser
  # Class for describing Document Background `w:background`
  class DocumentBackground
    attr_accessor :color1, :size, :color2, :image, :type

    def initialize(color1 = nil, type = 'simple')
      @color1 = color1
      @type = type
    end

    # Parse DocumentBackground object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [DocumentBackground] result of parsing
    def self.parse(node)
      background = DocumentBackground.new
      background.color1 = Color.from_int16(node.attribute('color').value)
      node.xpath('v:background').each do |second_color|
        unless second_color.attribute('targetscreensize').nil?
          background.size = second_color.attribute('targetscreensize').value.sub(',', 'x')
        end
        second_color.xpath('v:fill').each do |fill|
          if !fill.attribute('color2').nil?
            background.color2 = Color.from_int16(fill.attribute('color2').value.split(' ').first.delete('#'))
            background.type = fill.attribute('type').value
          elsif !fill.attribute('id').nil?
            path_to_media_file = OOXMLDocumentObject.get_link_from_rels(fill.attribute('id').value)
            path_to_image = OOXMLDocumentObject.copy_media_file("#{OOXMLDocumentObject.root_subfolder}/#{path_to_media_file.gsub('..', '')}")
            background.image = path_to_image
            background.type = 'image'
          end
        end
      end
      background
    end
  end
end

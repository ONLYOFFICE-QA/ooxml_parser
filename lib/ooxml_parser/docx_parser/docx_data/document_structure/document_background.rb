module OoxmlParser
  # Class for describing Document Background `w:background`
  class DocumentBackground < OOXMLDocumentObject
    attr_accessor :color1, :size, :color2, :type
    # @return [FileReference] image structure
    attr_accessor :file_reference

    def initialize(parent: nil)
      @color1 = nil
      @type = 'simple'
      @parent = parent
    end

    # Parse DocumentBackground object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [DocumentBackground] result of parsing
    def parse(node)
      @color1 = Color.new(parent: self).parse_hex_string(node.attribute('color').value)
      node.xpath('v:background').each do |second_color|
        unless second_color.attribute('targetscreensize').nil?
          @size = second_color.attribute('targetscreensize').value.sub(',', 'x')
        end
        second_color.xpath('v:fill').each do |fill|
          if !fill.attribute('color2').nil?
            @color2 = Color.new(parent: self).parse_hex_string(fill.attribute('color2').value.split(' ').first.delete('#'))
            @type = fill.attribute('type').value
          elsif !fill.attribute('id').nil?
            @file_reference = FileReference.new(parent: self).parse(fill)
            @type = 'image'
          end
        end
      end
      self
    end
  end
end

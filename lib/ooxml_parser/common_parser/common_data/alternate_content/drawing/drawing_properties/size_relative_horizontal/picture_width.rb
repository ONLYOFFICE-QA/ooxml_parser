module OoxmlParser
  # Class for parsing `wp14:pctWidth` object
  class PictureWidth
    # @return [Float] value of picture width
    attr_accessor :value

    # Parse PictureWidth
    # @param [Nokogiri::XML:Node] node with PictureWidth
    # @return [PictureWidth] result of parsing
    def self.parse(node)
      width = PictureWidth.new
      # value should be in percent as said in
      # https://msdn.microsoft.com/en-us/library/documentformat.openxml.office2010.word.drawing.relativewidth%28v=office.14%29.aspx?f=255&MSPPError=-2147217396
      width.value = node.child.to_s.to_f / 1000
      width
    end
  end
end

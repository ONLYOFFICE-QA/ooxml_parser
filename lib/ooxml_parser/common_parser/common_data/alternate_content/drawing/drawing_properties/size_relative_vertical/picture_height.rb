module OoxmlParser
  # Class for parsing `wp14:pctHeight` object
  class PictureHeight
    # @return [Float] value of picture width
    attr_accessor :value

    # Parse PictureHeight
    # @param [Nokogiri::XML:Node] node with PictureHeight
    # @return [PictureHeight] result of parsing
    def self.parse(node)
      width = PictureHeight.new
      # value should be in percent as said in
      # https://msdn.microsoft.com/en-us/library/documentformat.openxml.office2010.word.drawing.relativewidth%28v=office.14%29.aspx?f=255&MSPPError=-2147217396
      width.value = node.child.to_s.to_f / 1000
      width
    end
  end
end

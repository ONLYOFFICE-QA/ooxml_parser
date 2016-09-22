module OoxmlParser
  # Class for parsing `wp14:pctWidth`, `wp14:pctHeight` object
  class PictureDimension < OOXMLDocumentObject
    # @return [Float] value of picture width or height
    attr_accessor :value

    # Parse PictureWidth
    # @param [Nokogiri::XML:Node] node with PictureWidth
    # @return [PictureWidth] result of parsing
    def parse(node)
      # value should be in percent as said in
      # https://msdn.microsoft.com/en-us/library/documentformat.openxml.office2010.word.drawing.relativewidth%28v=office.14%29.aspx?f=255&MSPPError=-2147217396
      @value = OoxmlSize.new(node.child.text.to_f, :one_1000th_percent)
      self
    end
  end
end

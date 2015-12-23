module OoxmlParser
  class Stretching
    attr_accessor :fill_rectangle

    def self.parse(stretch_node)
      stretching = Stretching.new
      stretching.fill_rectangle = FillRectangle.parse(stretch_node.xpath('a:fillRect').first) if stretch_node.xpath('a:fillRect').first
      stretching
    end
  end
end

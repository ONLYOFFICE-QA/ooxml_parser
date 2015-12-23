# Image Data
module OoxmlParser
  class Image
    attr_accessor :path, :start_x, :start_y, :end_x, :end_y, :wrap_text

    def initialize
      @path = nil
      @start_x = 0
      @start_y = 0
      @end_x = nil
      @end_y = nil
      @wrap_text = 'In text'
    end
  end
end

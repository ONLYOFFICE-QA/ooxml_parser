module OoxmlParser
  class TransitionProperties
    attr_accessor :type, :through_black, :direction, :orientation, :spokes

    def initialize(type = nil)
      @type = type
    end
  end
end

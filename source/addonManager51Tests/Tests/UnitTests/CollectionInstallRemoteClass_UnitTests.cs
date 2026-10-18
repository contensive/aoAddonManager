using Xunit;
using Contensive.Processor;
using Contensive.AddonManager.Addons;

namespace Tests.UnitTests {
    /// <summary>
    /// Unit tests for CollectionInstallRemoteClass pure static methods
    /// </summary>
    public class CollectionInstallRemoteClass_UnitTests {

        [Fact]
        public void JsonEscape_NullInput_ReturnsEmptyString() {
            // Arrange
            string input = null;

            // Act
            string result = CollectionInstallRemoteClass.jsonEscape(input);

            // Assert
            Assert.Equal("", result);
        }

        [Fact]
        public void JsonEscape_EmptyString_ReturnsEmptyString() {
            // Arrange
            string input = "";

            // Act
            string result = CollectionInstallRemoteClass.jsonEscape(input);

            // Assert
            Assert.Equal("", result);
        }

        [Fact]
        public void JsonEscape_SpecialCharacters_EscapesCorrectly() {
            // Arrange
            string input = "Hello \"World\"\nNew Line\r\nWindows Line\tTab\\Backslash";

            // Act
            string result = CollectionInstallRemoteClass.jsonEscape(input);

            // Assert
            Assert.Contains("\\\"", result); // escaped quote
            Assert.Contains("\\n", result);  // escaped newline
            Assert.Contains("\\r", result);  // escaped carriage return
            Assert.Contains("\\t", result);  // escaped tab
            Assert.Contains("\\\\", result); // escaped backslash
        }

        [Fact]
        public void JsonEscape_PlainText_ReturnsUnchanged() {
            // Arrange
            string input = "Simple text without special characters";

            // Act
            string result = CollectionInstallRemoteClass.jsonEscape(input);

            // Assert
            Assert.Equal(input, result);
        }

        [Theory]
        [InlineData(true, "true")]
        [InlineData(false, "false")]
        public void BoolJson_ValidBoolean_ReturnsCorrectString(bool input, string expected) {
            // Act
            string result = CollectionInstallRemoteClass.boolJson(input);

            // Assert
            Assert.Equal(expected, result);
        }

        [Fact]
        public void JsonError_SimpleMessage_ReturnsValidJson() {
            // Arrange
            string errorMessage = "Test error message";

            // Act
            string result = CollectionInstallRemoteClass.jsonError(errorMessage);

            // Assert
            Assert.StartsWith("{", result);
            Assert.EndsWith("}", result);
            Assert.Contains("\"error\":true", result);
            Assert.Contains("Test error message", result);
        }

        [Fact]
        public void JsonError_MessageWithSpecialCharacters_EscapesCorrectly() {
            // Arrange
            string errorMessage = "Error with \"quotes\" and\nnewlines";

            // Act
            string result = CollectionInstallRemoteClass.jsonError(errorMessage);

            // Assert
            Assert.StartsWith("{", result);
            Assert.EndsWith("}", result);
            Assert.Contains("\"error\":true", result);
            Assert.Contains("\\\"", result); // escaped quotes
            Assert.Contains("\\n", result);  // escaped newline
        }

        [Fact]
        public void JsonError_EmptyMessage_ReturnsValidJson() {
            // Arrange
            string errorMessage = "";

            // Act
            string result = CollectionInstallRemoteClass.jsonError(errorMessage);

            // Assert
            Assert.StartsWith("{", result);
            Assert.EndsWith("}", result);
            Assert.Contains("\"error\":true", result);
        }

        [Fact]
        public void ExtractCollectionGuid_ValidGuid_ReturnsGuid() {
            // Arrange
            string validGuid = "{12345678-1234-1234-1234-123456789012}";

            // Act
            string result = CollectionInstallRemoteClass.extractCollectionGuid(validGuid);

            // Assert
            Assert.Equal(validGuid, result);
        }

        [Fact]
        public void ExtractCollectionGuid_GuidWithoutBraces_ReturnsSameString() {
            // Arrange
            string guidWithoutBraces = "12345678-1234-1234-1234-123456789012";

            // Act
            string result = CollectionInstallRemoteClass.extractCollectionGuid(guidWithoutBraces);

            // Assert
            Assert.Equal(guidWithoutBraces, result);
        }

        [Fact]
        public void ExtractCollectionGuid_EmptyString_ReturnsEmptyString() {
            // Arrange
            string input = "";

            // Act
            string result = CollectionInstallRemoteClass.extractCollectionGuid(input);

            // Assert
            Assert.Equal("", result);
        }

        [Fact]
        public void ExtractCollectionGuid_NullInput_ReturnsEmptyString() {
            // Arrange
            string input = null;

            // Act
            string result = CollectionInstallRemoteClass.extractCollectionGuid(input);

            // Assert
            Assert.Equal("", result);
        }

        [Fact]
        public void ExtractCollectionGuid_InvalidFormat_ReturnsOriginalString() {
            // Arrange
            string invalidInput = "not-a-guid";

            // Act
            string result = CollectionInstallRemoteClass.extractCollectionGuid(invalidInput);

            // Assert
            Assert.Equal(invalidInput, result);
        }
    }
}

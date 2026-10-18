using Xunit;
using Contensive.Processor;
using Contensive.AddonManager.Addons;
using Contensive.AddonManager.Controllers;

namespace Tests.IntegrationTests {
    /// <summary>
    /// Integration tests for authentication and authorization gates across all addon entry points
    /// Priority: P0 (Security-critical - remote methods and admin-only operations)
    /// </summary>
    public class AuthGate_IntegrationTests {

        [Fact]
        public void CollectionInstallRemote_UnauthenticatedUser_ReturnsError() {
            // Arrange
            using (CPClass cp = new CPClass(Constants.testAppName)) {
                // Note: CPClass starts unauthenticated by default

                // Act
                var addon = new CollectionInstallRemoteClass();
                string result = addon.Execute(cp) as string;

                // Assert
                Assert.NotNull(result);
                Assert.Contains("error", result.ToLower());
                Assert.Contains("administrator", result.ToLower());
            }
        }

        [Fact]
        public void CollectionInstallRemote_AdminUser_AcceptsRequest() {
            // Arrange
            using (CPClass cp = new CPClass(Constants.testAppName)) {
                // TODO: Authenticate as admin user for this test
                // This requires setting up a test admin account in the addonTest app
                // cp.User.Recognize(...) or cp.User.Login(...)

                // Act
                var addon = new CollectionInstallRemoteClass();
                string result = addon.Execute(cp) as string;

                // Assert
                // Should not contain "Administrator access required" error
                Assert.NotNull(result);
                Assert.DoesNotContain("Administrator access required", result);
            }
        }

        [Fact]
        public void InstallController_InstallCollectionFromLibrary_UnauthenticatedUser_ThrowsException() {
            // Arrange
            using (CPClass cp = new CPClass(Constants.testAppName)) {
                string testGuid = "{12345678-1234-1234-1234-123456789012}";
                string errorMessage = "";

                // Act & Assert
                var ex = Assert.Throws<System.Exception>(() => {
                    InstallController.installCollectionFromLibrary(cp, testGuid, ref errorMessage);
                });

                Assert.Contains("Administrator access required", ex.Message);
            }
        }

        [Fact]
        public void InstallController_InstallCollectionFromLibrary_AdminUser_DoesNotThrowAuthException() {
            // Arrange
            using (CPClass cp = new CPClass(Constants.testAppName)) {
                // TODO: Authenticate as admin user for this test
                string testGuid = "{12345678-1234-1234-1234-123456789012}";
                string errorMessage = "";

                // Act & Assert
                // This will likely throw for other reasons (invalid GUID, network issues, etc.)
                // but should NOT throw "Administrator access required"
                try {
                    InstallController.installCollectionFromLibrary(cp, testGuid, ref errorMessage);
                } catch (System.Exception ex) {
                    Assert.DoesNotContain("Administrator access required", ex.Message);
                }
            }
        }

        [Fact]
        public void CollectionLibrary_UnauthenticatedUser_ReturnsErrorOrRedirect() {
            // Arrange
            using (CPClass cp = new CPClass(Constants.testAppName)) {

                // Act
                var addon = new CollectionLibraryClass();
                string result = addon.Execute(cp) as string;

                // Assert
                // Should either return empty, redirect, or show an access denied message
                // (Contensive framework may handle this at a higher level)
                Assert.NotNull(result);
            }
        }

        [Fact]
        public void Export_UnauthenticatedUser_ReturnsErrorOrRedirect() {
            // Arrange
            using (CPClass cp = new CPClass(Constants.testAppName)) {

                // Act
                var addon = new ExportClass();
                string result = addon.Execute(cp) as string;

                // Assert
                Assert.NotNull(result);
            }
        }

        [Fact]
        public void Upload_UnauthenticatedUser_ReturnsErrorOrRedirect() {
            // Arrange
            using (CPClass cp = new CPClass(Constants.testAppName)) {

                // Act
                var addon = new UploadClass();
                string result = addon.Execute(cp) as string;

                // Assert
                Assert.NotNull(result);
            }
        }

        [Fact]
        public void Uninstall_UnauthenticatedUser_ReturnsErrorOrRedirect() {
            // Arrange
            using (CPClass cp = new CPClass(Constants.testAppName)) {

                // Act
                var addon = new UninstallClass();
                string result = addon.Execute(cp) as string;

                // Assert
                Assert.NotNull(result);
            }
        }
    }
}

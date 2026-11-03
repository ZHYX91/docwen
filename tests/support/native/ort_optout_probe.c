/* Linux opt-out verification only. Blocks and records libc network calls;
 * deliberately does not claim coverage of direct syscalls or static binaries. */
#define _GNU_SOURCE
#include <dlfcn.h>
#include <errno.h>
#include <fcntl.h>
#include <netdb.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <sys/socket.h>
#include <unistd.h>

static void record(const char *event) {
    const char *path = getenv("DOCWEN_NATIVE_PROBE_LOG");
    if (!path) return;
    int fd = open(path, O_WRONLY | O_CREAT | O_APPEND, 0600);
    if (fd >= 0) { write(fd, event, strlen(event)); write(fd, "\n", 1); close(fd); }
}
__attribute__((constructor)) static void loaded(void) { record("probe-loaded"); }
static int missing_vendor(const char *path) {
    if (strcmp(path, "/sys/class/drm/card0/device/vendor") != 0) return 0;
    record("vendor-ENOENT"); errno = ENOENT; return 1;
}
FILE *fopen(const char *path, const char *mode) {
    if (missing_vendor(path)) return NULL;
    FILE *(*real)(const char *, const char *) = dlsym(RTLD_NEXT, "fopen");
    return real(path, mode);
}
FILE *fopen64(const char *path, const char *mode) {
    if (missing_vendor(path)) return NULL;
    FILE *(*real)(const char *, const char *) = dlsym(RTLD_NEXT, "fopen64");
    return real(path, mode);
}
int socket(int domain, int type, int protocol) {
    (void)domain; (void)type; (void)protocol;
    record("network:socket"); errno = EPERM; return -1;
}
int connect(int fd, const struct sockaddr *address, socklen_t length) {
    (void)fd; (void)address; (void)length;
    record("network:connect"); errno = EPERM; return -1;
}
ssize_t send(int fd, const void *buffer, size_t size, int flags) {
    (void)fd; (void)buffer; (void)size; (void)flags;
    record("network:send"); errno = EPERM; return -1;
}
ssize_t sendto(int fd, const void *buffer, size_t size, int flags,
               const struct sockaddr *address, socklen_t length) {
    (void)fd; (void)buffer; (void)size; (void)flags; (void)address; (void)length;
    record("network:sendto"); errno = EPERM; return -1;
}
ssize_t sendmsg(int fd, const struct msghdr *message, int flags) {
    (void)fd; (void)message; (void)flags;
    record("network:sendmsg"); errno = EPERM; return -1;
}
int getaddrinfo(const char *node, const char *service,
                const struct addrinfo *hints, struct addrinfo **result) {
    (void)node; (void)service; (void)hints; (void)result;
    record("network:getaddrinfo"); return EAI_FAIL;
}
